from odoo import http
from odoo.http import request
from odoo.exceptions import ValidationError
from jinja2 import Environment, FileSystemLoader
from functools import lru_cache
import json
import html as html_lib
import os
import re
import logging
import time
from urllib.parse import quote_plus

_logger = logging.getLogger(__name__)


@lru_cache(maxsize=8)
def _get_cached_template_env(templates_path):
    env = Environment(loader=FileSystemLoader(templates_path))
    default_format = env.filters.get('format')

    def _fallback_format(fmt, *args, **kwargs):
        if default_format:
            return default_format(fmt, *args, **kwargs)
        return fmt % args

    def format_filter(fmt, *args, **kwargs):
        # Ensure consistent accounting formatting:
        # - thousands separators (12,345.67)
        # - negatives in parentheses ((12,345.67))
        # Applies to printf-style float formats like "%.2f", "%.0f", etc.
        if isinstance(fmt, str) and args:
            match = re.fullmatch(r"%\.(\d+)f", fmt)
        else:
            match = None
        if match and args:
            decimals = 0
            value = args[0] or 0.0
            try:
                number = float(value)
            except (TypeError, ValueError):
                return _fallback_format(fmt, *args, **kwargs)
            rounded_number = round(number, decimals)
            if rounded_number == 0:
                return "-"
            if rounded_number < 0:
                return f"({abs(rounded_number):,.{decimals}f})"
            return f"{rounded_number:,.{decimals}f}"
        return _fallback_format(fmt, *args, **kwargs)

    def format_percent(fmt, *args, **kwargs):
        """Format percentages using the exact precision requested by fmt."""
        if isinstance(fmt, str) and args:
            match = re.fullmatch(r"%\.(\d+)f", fmt)
        else:
            match = None
        if match and args:
            decimals = int(match.group(1))
            value = args[0] or 0.0
            try:
                number = float(value)
            except (TypeError, ValueError):
                return _fallback_format(fmt, *args, **kwargs)
            formatted = f"{abs(number):,.{decimals}f}"
            if number < 0:
                return f"({formatted})"
            return formatted
        return _fallback_format(fmt, *args, **kwargs)

    def running_case_filter(value):
        """Convert headings/labels to sentence case while keeping acronyms (VAT/IFRS/UAE/ROU/etc)."""
        if value is None:
            return ''
        text = str(value).strip()
        if not text:
            return ''
        # Normalize ampersands in labels (e.g., "Revenue & other income" -> "Revenue and other income").
        text = re.sub(r'\s*&\s*', ' and ', text)
        text = re.sub(r'\s{2,}', ' ', text).strip()

        parts = re.split(r'(\s+)', text)
        out = []
        for part in parts:
            if not part or part.isspace():
                out.append(part)
                continue
            token = part
            # Keep acronyms and codes (all-caps), and anything containing digits.
            if token.isupper() or any(ch.isdigit() for ch in token):
                out.append(token)
            else:
                out.append(token.lower())

        # Capitalize the first alphabetic character (if the first token isn't an acronym).
        for i, part in enumerate(out):
            if not part or part.isspace():
                continue
            if part.isupper():  # acronym in first position
                break
            for j, ch in enumerate(part):
                if ch.isalpha():
                    out[i] = part[:j] + ch.upper() + part[j + 1:]
                    break
            break
        return ''.join(out)

    env.filters['format'] = format_filter
    env.filters['format_percent'] = format_percent
    env.filters['rcase'] = running_case_filter
    return env


@lru_cache(maxsize=32)
def _get_cached_text_file(path, mtime):
    with open(path, 'r', encoding='utf-8') as file_obj:
        return file_obj.read()


class AuditReportController(http.Controller):
    _DMCC_FREEZONE_NAME = 'Dubai Multi Commodities Centre Free Zone'

    def _is_dmcc_summary_enabled(self, report):
        return (getattr(report.company_id, 'free_zone', '') or '') == self._DMCC_FREEZONE_NAME

    def _module_path(self):
        return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    def _templates_path(self):
        return os.path.join(self._module_path(), 'templates')

    def _css_path(self):
        return os.path.join(self._templates_path(), 'audit_report_style.css')

    def _get_template_env(self, templates_path=None):
        return _get_cached_template_env(templates_path or self._templates_path())

    def _get_cached_css_content(self, css_path=None):
        path = css_path or self._css_path()
        try:
            mtime = os.path.getmtime(path)
        except OSError:
            return ''
        try:
            return _get_cached_text_file(path, mtime)
        except OSError:
            return ''

    def _html_to_pdf(self, html_content, base_url=None):
        """Render HTML+CSS to PDF server-side.

        WeasyPrint supports CSS paged media (@page, counters, margin boxes) and
        matches browser print output better than wkhtmltopdf in most cases.
        """
        try:
            from weasyprint import HTML
        except Exception as e:
            raise RuntimeError(f"WeasyPrint is not available in this Odoo environment: {e}")

        return HTML(string=html_content, base_url=base_url).write_pdf()

    def _html_to_pdf_doc(self, html_content, base_url=None):
        """Render HTML to a WeasyPrint document for page counting."""
        try:
            from weasyprint import HTML
        except Exception as e:
            raise RuntimeError(f"WeasyPrint is not available in this Odoo environment: {e}")

        return HTML(string=html_content, base_url=base_url).render()

    def _default_toc_entries(self, report, data):
        signature_names = data.get('signature_names') or []
        director_label_lower = 'director' if len(signature_names) == 1 else 'directors'
        entries = [
            {'label': 'Entity information', 'page_range': '1'},
            {'label': f'Report of {director_label_lower}', 'page_range': '2-3'},
            {'label': 'Independent auditor report', 'page_range': '4-5'},
            {'label': 'Statement of financial position', 'page_range': '6'},
            {'label': 'Statement of profit and loss and other comprehensive income', 'page_range': '7'},
            {'label': 'Statement of changes in equity', 'page_range': '8'},
            {'label': 'Statement of cash flows', 'page_range': '9'},
            {'label': 'Notes to the financial statements', 'page_range': '10-22'},
        ]
        if self._is_dmcc_summary_enabled(report):
            entries.append({'label': 'Summary sheet', 'page_range': '23'})
        return entries

    def _compute_toc_entries(self, report, report_data=None, template_env=None, css_content=None):
        module_path = self._module_path()
        data = report_data if report_data is not None else report._get_report_data()
        templates_path = self._templates_path()
        if template_env is None:
            template_env = self._get_template_env(templates_path)
        if css_content is None:
            css_content = self._get_cached_css_content(os.path.join(templates_path, 'audit_report_style.css'))
        signature_names = data.get('signature_names') or []
        director_label_lower = 'director' if len(signature_names) == 1 else 'directors'
        sections = [
            ('entity_information', 'Entity information'),
            ('report_of_directors', f'Report of {director_label_lower}'),
            ('independent_auditor_report', 'Independent auditor report'),
            ('balance_sheet_page', 'Statement of financial position'),
            ('profit_loss', 'Statement of profit and loss and other comprehensive income'),
            ('changes_in_equity', 'Statement of changes in equity'),
            ('cash_flows', 'Statement of cash flows'),
            ('notes_to_financial_statements', 'Notes to the financial statements'),
        ]
        if self._is_dmcc_summary_enabled(report):
            sections.append(('dmcc_sheet', 'Summary sheet'))

        default_toc_entries = self._default_toc_entries(report, data)
        toc_entries = []
        current_page = 1
        for section_key, label in sections:
            html_with_style = self._render_report_html(
                report,
                sections_to_render=[section_key],
                toc_entries=default_toc_entries,
                report_data=data,
                template_env=template_env,
                css_content=css_content,
            )
            doc = self._html_to_pdf_doc(html_with_style, base_url=module_path)
            page_count = len(doc.pages)
            if page_count <= 0:
                continue
            end_page = current_page + page_count - 1
            page_range = str(current_page) if page_count == 1 else f"{current_page}-{end_page}"
            toc_entries.append({'label': label, 'page_range': page_range})
            current_page = end_page + 1

        return toc_entries

    def _render_report_html(
        self,
        report,
        sections_to_render=None,
        toc_entries=None,
        report_data=None,
        template_env=None,
        css_content=None,
    ):
        company = report.company_id
        if not report.exists():
            return None

        templates_path = self._templates_path()
        env = template_env or self._get_template_env(templates_path)
        data = report_data if report_data is not None else report._get_report_data()
        period_category = (report.audit_period_category or '').lower()
        if period_category.startswith('dormant_'):
            template_name = (
                'audit_report_template_dormant_2y.html'
                if data.get('show_prior_year')
                else 'audit_report_template_dormant_1y.html'
            )
        elif period_category.startswith('cessation_'):
            template_name = (
                'audit_report_template_cessation_2y.html'
                if data.get('show_prior_year')
                else 'audit_report_template_cessation_1y.html'
            )
        else:
            template_name = (
                'audit_report_template_2y.html'
                if data.get('show_prior_year')
                else 'audit_report_template.html'
            )
        template = env.get_template(template_name)
        report_type_label = dict(report._fields['report_type'].selection).get(report.report_type, '')
        report_period_end = report.date_end.strftime("%d %B %Y") if report.date_end else ''
        report_title = f"{report_type_label} {report_period_end}".strip()
        report_ended_label = 'Year Ended' if report.report_type == 'year' else 'Period Ended'
        report_period_word = 'year' if report.report_type == 'year' else 'period'
        activity = getattr(company, 'trade_license_activities', '') or ''
        implementing_regulations_freezone = getattr(company, 'implementing_regulations_freezone', '') or ''
        if not implementing_regulations_freezone and hasattr(company, '_get_free_zone_implementing_regulations'):
            implementing_regulations_freezone = (
                company._get_free_zone_implementing_regulations(getattr(company, 'free_zone', '')) or ''
            )

        context = {
            'company': company,
            'company_name': getattr(company, 'name', ''),
            'freezone': getattr(company, 'free_zone', ''),
            'implementing_regulations_freezone': implementing_regulations_freezone,
            'license': getattr(company, 'company_license_number', ''),
            'owner': getattr(company, 'owner', ''),
            'business_activity': activity.lower(),
            'date_start': report.date_start,
            'date_end': report.date_end,
            'freezone_selection': report.auditor_type,
            'report_title': report_title,
            'report_type': report.report_type,
            'report_period_end': report_period_end,
            'report_ended_label': report_ended_label,
            'report_period_word': report_period_word,
            'sections_to_render': sections_to_render,
        }
        context.update(data)
        if toc_entries is None:
            toc_entries = self._default_toc_entries(report, data)
        context['toc_entries'] = toc_entries

        html_content = template.render(**context)
        if css_content is None:
            css_content = self._get_cached_css_content(os.path.join(templates_path, 'audit_report_style.css'))
        if not css_content:
            return html_content

        return html_content.replace('</head>', f'<style>{css_content}</style></head>')
    
    @http.route('/audit_report/view/<int:report_id>', type='http', auth='user')
    def view_report(self, report_id, **kwargs):
        report = request.env['audit.report'].browse(report_id)
        html_with_style = self._render_report_html(report)
        if not html_with_style:
            return request.not_found()
        
        return request.make_response(
            html_with_style,
            headers=[
                ('Content-Type', 'text/html; charset=utf-8'),
                ('Cache-Control', 'no-store, no-cache, must-revalidate, max-age=0'),
                ('Pragma', 'no-cache'),
                ('Expires', '0'),
            ]
        )

    @http.route('/audit_report/pdf/<int:report_id>', type='http', auth='user')
    def view_report_pdf(self, report_id, **kwargs):
        route_started = time.perf_counter()
        report = request.env['audit.report'].browse(report_id)
        report_data_started = time.perf_counter()
        report_data = report._get_report_data()
        report_data_elapsed_ms = (time.perf_counter() - report_data_started) * 1000.0
        _logger.debug(
            "AUDIT_PERF view_report_pdf report_id=%s stage=report_data elapsed_ms=%.2f",
            report.id,
            report_data_elapsed_ms,
        )
        templates_path = self._templates_path()
        template_env = self._get_template_env(templates_path)
        css_content = self._get_cached_css_content(os.path.join(templates_path, 'audit_report_style.css'))

        toc_started = time.perf_counter()
        try:
            toc_entries = self._compute_toc_entries(
                report,
                report_data=report_data,
                template_env=template_env,
                css_content=css_content,
            )
        except Exception as e:
            _logger.exception("Audit report TOC generation failed: %s", e)
            toc_entries = None
        toc_elapsed_ms = (time.perf_counter() - toc_started) * 1000.0
        _logger.debug(
            "AUDIT_PERF view_report_pdf report_id=%s stage=toc elapsed_ms=%.2f",
            report.id,
            toc_elapsed_ms,
        )

        html_started = time.perf_counter()
        html_with_style = self._render_report_html(
            report,
            toc_entries=toc_entries,
            report_data=report_data,
            template_env=template_env,
            css_content=css_content,
        )
        html_elapsed_ms = (time.perf_counter() - html_started) * 1000.0
        _logger.debug(
            "AUDIT_PERF view_report_pdf report_id=%s stage=html elapsed_ms=%.2f",
            report.id,
            html_elapsed_ms,
        )
        if not html_with_style:
            return request.not_found()

        pdf_started = time.perf_counter()
        try:
            module_path = self._module_path()
            pdf_content = self._html_to_pdf(html_with_style, base_url=module_path)
        except Exception as e:
            _logger.exception("Audit report PDF generation failed: %s", e)
            summary, issues = self._extract_error_summary_and_issues(
                e,
                'Unable to generate the audit report PDF.',
            )
            return self._render_editor_error_response(
                'Audit Report PDF Error',
                summary,
                issues,
                back_url=f'/web#id={report.id}&model=audit.report&view_type=form',
                status=500,
            )
        pdf_elapsed_ms = (time.perf_counter() - pdf_started) * 1000.0
        total_elapsed_ms = (time.perf_counter() - route_started) * 1000.0
        _logger.debug(
            "AUDIT_PERF view_report_pdf report_id=%s stage=pdf elapsed_ms=%.2f total_elapsed_ms=%.2f",
            report.id,
            pdf_elapsed_ms,
            total_elapsed_ms,
        )
        filename = f'Audit_Report_{report.id}.pdf'
        return request.make_response(
            pdf_content,
            headers=[
                ('Content-Type', 'application/pdf'),
                ('Content-Disposition', f'inline; filename="{filename}"'),
                ('Content-Length', str(len(pdf_content))),
                ('Cache-Control', 'no-store, no-cache, must-revalidate, max-age=0'),
                ('Pragma', 'no-cache'),
                ('Expires', '0'),
            ]
        )

    def _forbidden_response(self, message='Forbidden'):
        return request.make_response(
            message,
            headers=[('Content-Type', 'text/plain; charset=utf-8')],
            status=403,
        )

    def _load_revision(self, revision_id, include_removed=True):
        revision = request.env['audit.report.revision'].browse(revision_id)
        if not revision.exists():
            return None
        user_company_ids = request.env.user.company_ids.ids
        if revision.company_id.id not in user_company_ids:
            return None
        if not include_removed and revision.is_removed:
            return None
        return revision

    def _render_editor_shell(self, title, body_html):
        title_escaped = html_lib.escape(title or 'Editor')
        return f"""
<!doctype html>
<html>
  <head>
    <meta charset="utf-8" />
    <title>{title_escaped}</title>
    <style>
      body {{
        margin: 0;
        font-family: "Segoe UI", Tahoma, Arial, sans-serif;
        background: #e7e9ee;
        color: #1d2733;
      }}
      .page {{
        max-width: 1560px;
        margin: 0 auto;
        padding: 24px;
      }}
      .panel {{
        background: #fff;
        border: 1px solid #d5dbe6;
        border-radius: 10px;
        box-shadow: 0 4px 14px rgba(17, 24, 39, 0.05);
        padding: 16px;
        margin-bottom: 16px;
      }}
      .muted {{
        color: #6b7280;
        font-size: 13px;
      }}
      .ok {{
        background: #e8f7ec;
        border: 1px solid #9bd7ac;
        color: #0f5132;
        border-radius: 6px;
        padding: 10px 12px;
        margin-bottom: 12px;
      }}
      .error-panel {{
        background: #fff4f4;
        border: 1px solid #efb3b3;
        color: #7f1d1d;
      }}
      .error-list {{
        margin: 8px 0 0 20px;
        padding: 0;
      }}
      .error-list li {{
        margin: 0 0 4px 0;
      }}
      .btn {{
        display: inline-block;
        background: #1f5eff;
        color: #fff;
        border: none;
        border-radius: 6px;
        padding: 10px 14px;
        text-decoration: none;
        cursor: pointer;
        font-size: 14px;
      }}
      .btn-secondary {{
        background: #4b5563;
      }}
      .btn-danger {{
        background: #b91c1c;
      }}
      .links a {{
        margin-right: 8px;
      }}
      table {{
        border-collapse: collapse;
        width: 100%;
      }}
      th, td {{
        border: 1px solid #d0d7e2;
        padding: 6px;
        vertical-align: top;
      }}
      input[type=\"text\"], textarea, select {{
        width: 100%;
        box-sizing: border-box;
        padding: 8px;
        border: 1px solid #c6cdd9;
        border-radius: 6px;
        font-size: 14px;
      }}
      textarea {{
        min-height: 65vh;
        font-family: monospace;
      }}
      .table-list {{
        display: grid;
        grid-template-columns: 320px 1fr;
        gap: 16px;
      }}
      .table-list ul {{
        list-style: none;
        margin: 0;
        padding: 0;
      }}
      .table-list li {{
        border: 1px solid #d9dee8;
        border-radius: 6px;
        margin-bottom: 8px;
        background: #fff;
      }}
      .table-list li a {{
        display: block;
        padding: 10px 12px;
        text-decoration: none;
        color: #1d2733;
      }}
      .table-list li.active {{
        background: #ecf2ff;
        border-color: #9cb6ff;
      }}
      .small {{
        font-size: 12px;
        color: #6b7280;
      }}
      .toolbar {{
        display: flex;
        gap: 8px;
        flex-wrap: wrap;
        margin-bottom: 12px;
        background: #f6f8fc;
        border: 1px solid #d5dbea;
        border-radius: 8px;
        padding: 10px;
      }}
      .toolbar button {{
        background: #eef2ff;
        border: 1px solid #bfccff;
        color: #243b85;
        border-radius: 6px;
        padding: 8px 10px;
        cursor: pointer;
      }}
      .toolbar select,
      .toolbar input[type="color"] {{
        width: auto;
        min-width: 82px;
        padding: 6px 8px;
      }}
      .toolbar .toolbar-label {{
        font-size: 12px;
        color: #4b5563;
        display: inline-flex;
        align-items: center;
        gap: 6px;
      }}
      .visual-toolbar {{
        position: sticky;
        top: 10px;
        z-index: 20;
      }}
      .editor-stage {{
        background: #cdd3dc;
        border: 1px solid #b4becb;
        border-radius: 10px;
        padding: 18px;
        overflow-x: auto;
        overflow-y: visible;
        max-height: none;
      }}
      .editor-canvas {{
        width: 980px;
        max-width: 100%;
        margin: 0 auto;
        background: #fff;
        box-shadow: 0 14px 34px rgba(15, 23, 42, 0.35);
        border: 1px solid #c7d0dd;
        transform-origin: top center;
      }}
      .editor-frame {{
        width: 100%;
        height: 1280px;
        border: none;
        display: block;
        background: white;
      }}
      .editor-caption {{
        margin-top: 8px;
        color: #5a6474;
        font-size: 12px;
      }}
    </style>
  </head>
  <body>
    <div class=\"page\">
      {body_html}
    </div>
  </body>
</html>
"""

    def _extract_error_summary_and_issues(self, error, default_summary):
        raw = str(error or '').replace('\r', '\n')
        lines = [line.strip() for line in raw.split('\n') if line.strip()]
        if not lines:
            return default_summary, []

        summary = lines[0]
        issues = []
        for line in lines[1:]:
            cleaned = line.lstrip('-*• ').strip()
            if cleaned:
                issues.append(cleaned)

        if summary.startswith(('-', '*', '•')):
            cleaned_summary = summary.lstrip('-*• ').strip()
            summary = default_summary
            if cleaned_summary:
                issues.insert(0, cleaned_summary)

        return summary or default_summary, issues

    def _render_editor_error_response(self, title, summary, issues=None, back_url=None, status=400):
        safe_title = html_lib.escape(title or 'Error')
        safe_summary = html_lib.escape(summary or 'Something went wrong.')
        issue_items = ''.join(f"<li>{html_lib.escape(issue)}</li>" for issue in (issues or []))
        issue_list = f'<ul class="error-list">{issue_items}</ul>' if issue_items else ''
        back_link = (
            f'<a class="btn btn-secondary" href="{html_lib.escape(back_url)}">Back</a>'
            if back_url else ''
        )
        body = f"""
<div class="panel error-panel">
  <h2>{safe_title}</h2>
  <p>{safe_summary}</p>
  {issue_list}
  <br/>
  <div class="links">{back_link}</div>
</div>
"""
        page = self._render_editor_shell(title, body)
        return request.make_response(
            page,
            headers=[('Content-Type', 'text/html; charset=utf-8')],
            status=status,
        )

    def _persist_revision_edit(self, revision, prepared_html, force_new_revision=False):
        """Persist edited HTML with session-style revision behavior."""
        current_hash = revision._compute_html_hash()
        prepared_hash = revision._compute_html_hash(prepared_html)
        if prepared_hash == current_hash:
            return revision, 'unchanged'

        if force_new_revision:
            return revision.create_next_revision(prepared_html), 'new'

        # Session behavior:
        # - first save from a root revision creates one child revision
        # - subsequent saves on that child update it in place
        if revision.parent_revision_id:
            revision.write({'html_content': prepared_html})
            return revision, 'updated'

        return revision.create_next_revision(prepared_html), 'new'

    def _render_structured_editor_page(self, revision, selected_table_key=None, saved_state=''):
        tables = revision.get_table_index()
        if not tables:
            content = (
                "<div class='panel'><h2>Structured Table Editor</h2>"
                "<p class='muted'>No tables were found in this revision.</p></div>"
            )
            return self._render_editor_shell('Structured Table Editor', content)

        selected_table_key = selected_table_key or tables[0].get('key')
        if selected_table_key not in [table.get('key') for table in tables]:
            selected_table_key = tables[0].get('key')

        payload = revision.get_table_payload(selected_table_key)
        selected_table_meta = next(
            (table for table in tables if table.get('key') == selected_table_key),
            {}
        )
        is_soce_table = (selected_table_meta.get('section') or '') == 'changes_in_equity'
        has_nested_tables = bool(selected_table_meta.get('has_nested_tables'))
        csrf = request.csrf_token()
        if saved_state == 'new':
            saved_banner = "<div class='ok'>Saved as a new revision successfully.</div>"
        elif saved_state == 'updated':
            saved_banner = "<div class='ok'>Saved to the current revision session.</div>"
        elif saved_state == 'unchanged':
            saved_banner = "<div class='ok'>No changes detected; revision kept as-is.</div>"
        else:
            saved_banner = ""
        back_url = f"/web#id={revision.id}&model=audit.report.revision&view_type=form"
        preview_url = f"/audit_report/revision/{revision.id}/preview"
        freeform_url = f"/audit_report/revision/{revision.id}/edit/freeform"
        table_rows = payload.get('rows') or []
        row_count = int(payload.get('body_row_count') or len(table_rows) or 0)
        col_count = int(selected_table_meta.get('cols') or 0) or max((len(row) for row in table_rows), default=0)
        default_target_row = row_count or 1
        default_target_col = col_count or 1
        row_controls_html = (
            f"<div class=\"toolbar\">"
            f"<label class=\"small\">Body Row #:"
            f"<input type=\"number\" name=\"target_row\" min=\"1\" value=\"{default_target_row}\" style=\"width: 90px; margin-left: 4px;\" />"
            f"</label>"
            f"<button class=\"btn btn-secondary\" type=\"submit\" name=\"table_action\" value=\"add_row_after\">Add Row After</button>"
            f"<button class=\"btn btn-secondary\" type=\"submit\" name=\"table_action\" value=\"remove_row\">Remove Row</button>"
            f"</div>"
        )
        if is_soce_table:
            row_controls_html = (
                row_controls_html
            )
            column_controls_html = (
                f"<div class=\"toolbar\">"
                f"<label class=\"small\">Column #:"
                f"<input type=\"number\" name=\"target_col\" min=\"2\" value=\"{max(default_target_col, 2)}\" style=\"width: 90px; margin-left: 4px;\" />"
                f"</label>"
                f"<button class=\"btn btn-secondary\" type=\"submit\" name=\"table_action\" value=\"add_col_after\">Add SOCE Column After</button>"
                f"<button class=\"btn btn-secondary\" type=\"submit\" name=\"table_action\" value=\"remove_col\">Remove SOCE Column</button>"
                f"</div>"
                "<p class='small'><strong>SOCE rule:</strong> Column #1 is description and cannot be removed.</p>"
            )
        else:
            column_controls_html = (
                f"<div class=\"toolbar\">"
                f"<label class=\"small\">Column #:"
                f"<input type=\"number\" name=\"target_col\" min=\"1\" value=\"{default_target_col}\" style=\"width: 90px; margin-left: 4px;\" />"
                f"</label>"
                f"<button class=\"btn btn-secondary\" type=\"submit\" name=\"table_action\" value=\"add_col_after\">Add Column After</button>"
                f"<button class=\"btn btn-secondary\" type=\"submit\" name=\"table_action\" value=\"remove_col\">Remove Column</button>"
                f"</div>"
            )
        if has_nested_tables:
            table_limit_notice = (
                "<p class='small'><strong>Composite table:</strong> nested-table cells are locked in structured mode "
                "to protect layout. Row/column operations are available.</p>"
            )
        else:
            table_limit_notice = ""

        table_links = []
        for table in tables:
            key = table.get('key') or ''
            section = html_lib.escape(table.get('section') or 'section')
            preview = html_lib.escape(table.get('preview') or '')
            table_class = html_lib.escape(table.get('table_class') or '')
            rows = table.get('rows') or 0
            cols = table.get('cols') or 0
            table_type_bits = []
            if table.get('has_nested_tables'):
                table_type_bits.append('composite')
            if table.get('is_nested_within_table'):
                table_type_bits.append('nested')
            table_type = html_lib.escape(', '.join(table_type_bits)) if table_type_bits else 'leaf'
            active_class = 'active' if key == selected_table_key else ''
            url = f"/audit_report/revision/{revision.id}/edit/structured?table_key={quote_plus(key)}"
            table_links.append(
                "<li class='{active}'>"
                "<a href='{url}'><strong>{section} / {key}</strong><br>"
                "<span class='small'>type: {table_type}</span><br>"
                "<span class='small'>class: {table_class}</span><br>"
                "<span class='small'>{rows} rows x {cols} cols</span><br>"
                "<span class='small'>{preview}</span></a></li>".format(
                    active=active_class,
                    url=url,
                    section=section,
                    key=html_lib.escape(key),
                    table_type=table_type,
                    table_class=table_class or '-',
                    rows=rows,
                    cols=cols,
                    preview=preview,
                )
            )

        matrix_rows = []
        for row in payload.get('rows') or []:
            cells = []
            for cell in row:
                value = html_lib.escape(cell.get('value') or '')
                field_name = html_lib.escape(cell.get('name') or '')
                tag = (cell.get('tag') or 'td').lower()
                editable = bool(cell.get('editable', True))
                readonly_attrs = ''
                if not editable:
                    readonly_attrs = (
                        " readonly title='Complex cell: use Freeform Editor for this content.'"
                    )
                if tag == 'th':
                    cells.append(
                        f"<th><input type='text' name='{field_name}' value='{value}'{readonly_attrs} /></th>"
                    )
                else:
                    cells.append(
                        f"<td><input type='text' name='{field_name}' value='{value}'{readonly_attrs} /></td>"
                    )
            matrix_rows.append(f"<tr>{''.join(cells)}</tr>")

        body = f"""
{saved_banner}
<div class=\"panel\">
  <h2>Structured Table Editor</h2>
  <p class=\"muted\">Revision v{revision.version_no} of {html_lib.escape(revision.document_id.name)}</p>
  <div class=\"links\">
    <a class=\"btn btn-secondary\" href=\"{back_url}\">Back to Revision</a>
    <a class=\"btn btn-secondary\" href=\"{preview_url}\" target=\"_blank\">Preview</a>
    <a class=\"btn\" href=\"{freeform_url}\" target=\"_blank\">Open Freeform Editor</a>
  </div>
</div>
<div class=\"table-list\">
  <div class=\"panel\">
    <h3>All Tables</h3>
    <p class=\"small\">Tables with nested content are listed. Nested-table cells are read-only in structured mode for safety.</p>
    <ul>{''.join(table_links)}</ul>
  </div>
  <div class=\"panel\">
    <h3>Editing: {html_lib.escape(selected_table_key)}</h3>
    <p class=\"muted\">You can edit values below, and also add/remove rows or columns.</p>
    {table_limit_notice}
    <form method=\"post\" action=\"/audit_report/revision/{revision.id}/save/structured\">
      <input type=\"hidden\" name=\"csrf_token\" value=\"{html_lib.escape(csrf)}\" />
      <input type=\"hidden\" name=\"table_key\" value=\"{html_lib.escape(selected_table_key)}\" />
      {row_controls_html}
      {column_controls_html}
      <table>{''.join(matrix_rows)}</table>
      <br/>
      <input type=\"hidden\" name=\"save_mode\" value=\"session\" />
      <button class=\"btn\" type=\"submit\" name=\"table_action\" value=\"save\" onclick=\"this.form.save_mode.value='session'\">Save Changes</button>
      <button class=\"btn btn-secondary\" type=\"submit\" name=\"table_action\" value=\"save\" onclick=\"this.form.save_mode.value='new'\">Save as New Revision</button>
    </form>
  </div>
</div>
"""
        return self._render_editor_shell('Structured Table Editor', body)

    def _render_freeform_editor_page(self, revision, saved_state=''):
        csrf = request.csrf_token()
        if saved_state == 'new':
            saved_banner = "<div class='ok'>Saved as a new revision successfully.</div>"
        elif saved_state == 'updated':
            saved_banner = "<div class='ok'>Saved to the current revision session.</div>"
        elif saved_state == 'unchanged':
            saved_banner = "<div class='ok'>No changes detected; revision kept as-is.</div>"
        else:
            saved_banner = ""
        initial_html_js = json.dumps(revision.html_content or '')
        back_url = f"/web#id={revision.id}&model=audit.report.revision&view_type=form"
        preview_url = f"/audit_report/revision/{revision.id}/preview"
        structured_url = f"/audit_report/revision/{revision.id}/edit/structured"

        body = f"""
{saved_banner}
<div class=\"panel\">
  <h2>Visual Layout Editor</h2>
  <p class=\"muted\">Revision v{revision.version_no} of {html_lib.escape(revision.document_id.name)}</p>
  <p class=\"muted\">Edit directly on the report layout below (non-technical mode). Use Save Changes for this session, or Save as New Revision for a checkpoint.</p>
  <div class=\"links\">
    <a class=\"btn btn-secondary\" href=\"{back_url}\">Back to Revision</a>
    <a class=\"btn btn-secondary\" href=\"{preview_url}\" target=\"_blank\">Preview</a>
    <a class=\"btn\" href=\"{structured_url}\" target=\"_blank\">Open Structured Editor</a>
  </div>
</div>
<div class=\"panel\">
  <div class=\"toolbar visual-toolbar\">
    <button type=\"button\" data-cmd=\"bold\">Bold</button>
    <button type=\"button\" data-cmd=\"italic\">Italic</button>
    <button type=\"button\" data-cmd=\"underline\">Underline</button>
    <button type=\"button\" data-cmd=\"insertUnorderedList\">Bullet List</button>
    <button type=\"button\" data-cmd=\"justifyLeft\">Align Left</button>
    <button type=\"button\" data-cmd=\"justifyCenter\">Align Center</button>
    <button type=\"button\" data-cmd=\"justifyRight\">Align Right</button>
    <button type=\"button\" id=\"undoBtn\">Undo</button>
    <button type=\"button\" id=\"redoBtn\">Redo</button>
    <button type=\"button\" id=\"resetBtn\">Reset</button>
    <label class=\"toolbar-label\">Zoom
      <select id=\"editorZoom\">
        <option value=\"70\">70%</option>
        <option value=\"85\">85%</option>
        <option value=\"100\" selected>100%</option>
        <option value=\"115\">115%</option>
        <option value=\"130\">130%</option>
      </select>
    </label>
    <label class=\"toolbar-label\">Border width
      <select id=\"borderWidthInput\">
        <option value=\"0.5px\">0.5px</option>
        <option value=\"1px\" selected>1px</option>
        <option value=\"1.5px\">1.5px</option>
        <option value=\"2px\">2px</option>
        <option value=\"3px\">3px</option>
      </select>
    </label>
    <label class=\"toolbar-label\">Border style
      <select id=\"borderStyleInput\">
        <option value=\"solid\" selected>Solid</option>
        <option value=\"dashed\">Dashed</option>
        <option value=\"dotted\">Dotted</option>
        <option value=\"double\">Double</option>
      </select>
    </label>
    <label class=\"toolbar-label\">Border color
      <input type=\"color\" id=\"borderColorInput\" value=\"#000000\" />
    </label>
    <button type=\"button\" data-border=\"all\">Border All</button>
    <button type=\"button\" data-border=\"top\">Border Top</button>
    <button type=\"button\" data-border=\"right\">Border Right</button>
    <button type=\"button\" data-border=\"bottom\">Border Bottom</button>
    <button type=\"button\" data-border=\"left\">Border Left</button>
    <button type=\"button\" data-border=\"clear\">Clear Border</button>
  </div>
  <p class=\"small\">Border actions apply to the selected table cell when inside a table, otherwise to the selected block element.</p>
  <div class=\"editor-stage\">
    <div class=\"editor-canvas\" id=\"editorCanvas\">
      <iframe id=\"visualEditorFrame\" class=\"editor-frame\"></iframe>
    </div>
  </div>
  <p class=\"editor-caption\">Canvas view emulates a document page for cleaner editing (Acrobat-style).</p>
  <form method=\"post\" action=\"/audit_report/revision/{revision.id}/save/freeform\">
    <input type=\"hidden\" name=\"csrf_token\" value=\"{html_lib.escape(csrf)}\" />
    <input type=\"hidden\" id=\"visualHtmlPayload\" name=\"html_content\" value=\"\" />
    <input type=\"hidden\" id=\"visualSaveMode\" name=\"save_mode\" value=\"session\" />
    <br/><br/>
    <button class=\"btn\" type=\"button\" id=\"saveVisualBtn\" data-save-mode=\"session\">Save Changes</button>
    <button class=\"btn btn-secondary\" type=\"button\" id=\"saveVisualNewBtn\" data-save-mode=\"new\">Save as New Revision</button>
  </form>
</div>
<script>
  (function () {{
    const initialHtml = {initial_html_js};
    const iframe = document.getElementById('visualEditorFrame');
    const editorCanvas = document.getElementById('editorCanvas');
    const payload = document.getElementById('visualHtmlPayload');
    const saveBtn = document.getElementById('saveVisualBtn');
    const saveNewBtn = document.getElementById('saveVisualNewBtn');
    const saveModeInput = document.getElementById('visualSaveMode');
    const resetBtn = document.getElementById('resetBtn');
    const undoBtn = document.getElementById('undoBtn');
    const redoBtn = document.getElementById('redoBtn');
    const commandButtons = Array.from(document.querySelectorAll('.visual-toolbar button[data-cmd]'));
    const borderButtons = Array.from(document.querySelectorAll('.visual-toolbar button[data-border]'));
    const zoomInput = document.getElementById('editorZoom');
    const borderWidthInput = document.getElementById('borderWidthInput');
    const borderStyleInput = document.getElementById('borderStyleInput');
    const borderColorInput = document.getElementById('borderColorInput');

    function loadEditorContent() {{
      iframe.srcdoc = initialHtml || '<!doctype html><html><head><meta charset=\"utf-8\"></head><body></body></html>';
    }}

    function applyCanvasZoom() {{
      if (!editorCanvas || !zoomInput) {{
        return;
      }}
      const zoomPct = parseInt(zoomInput.value || '100', 10);
      const scale = Math.max(40, Math.min(200, zoomPct)) / 100;
      editorCanvas.style.transform = 'scale(' + scale + ')';
    }}

    function focusEditor() {{
      if (iframe.contentWindow) {{
        iframe.contentWindow.focus();
      }}
    }}

    function normalizeAuditBullets(doc) {{
      if (!doc || !doc.body) {{
        return;
      }}
      const lists = Array.from(doc.body.querySelectorAll('ul'));
      lists.forEach(function (list) {{
        if (!list.classList.contains('audit-bullets')) {{
          list.classList.add('audit-bullets');
        }}
        if (!list.classList.contains('tight')) {{
          list.classList.add('tight');
        }}
      }});
    }}

    let frameResizeObserver = null;

    function resizeEditorFrame() {{
      const doc = iframe.contentDocument;
      if (!doc || !doc.documentElement) {{
        return;
      }}
      const body = doc.body;
      const html = doc.documentElement;
      const contentHeight = Math.max(
        body ? body.scrollHeight : 0,
        body ? body.offsetHeight : 0,
        html.scrollHeight || 0,
        html.offsetHeight || 0,
        1280
      );
      iframe.style.height = (contentHeight + 12) + 'px';
    }}

    iframe.addEventListener('load', function () {{
      const doc = iframe.contentDocument;
      if (!doc) {{
        return;
      }}
      try {{
        doc.designMode = 'on';
      }} catch (e) {{
        // Fallback if designMode is restricted.
      }}
      if (doc.body) {{
        doc.body.setAttribute('contenteditable', 'true');
      }}
      if (doc.documentElement) {{
        doc.documentElement.style.overflowY = 'hidden';
      }}
      if (doc.body) {{
        doc.body.style.overflowY = 'hidden';
      }}
      normalizeAuditBullets(doc);

      const triggerResize = function () {{
        if (window.requestAnimationFrame) {{
          window.requestAnimationFrame(resizeEditorFrame);
        }} else {{
          resizeEditorFrame();
        }}
      }};

      if (frameResizeObserver) {{
        frameResizeObserver.disconnect();
      }}
      if (window.MutationObserver && doc.body) {{
        frameResizeObserver = new MutationObserver(triggerResize);
        frameResizeObserver.observe(doc.body, {{
          subtree: true,
          childList: true,
          characterData: true,
          attributes: true,
        }});
      }}

      doc.addEventListener('input', triggerResize, true);
      doc.addEventListener('keyup', triggerResize, true);
      triggerResize();
    }});

    commandButtons.forEach(function (button) {{
      button.addEventListener('click', function () {{
        const cmd = button.getAttribute('data-cmd');
        const doc = iframe.contentDocument;
        if (!doc) {{
          return;
        }}
        focusEditor();
        doc.execCommand(cmd, false, null);
        if (cmd === 'insertUnorderedList') {{
          normalizeAuditBullets(doc);
        }}
      }});
    }});

    function getSelectionElement(doc) {{
      const selection = doc.getSelection ? doc.getSelection() : null;
      if (selection && selection.rangeCount > 0) {{
        let node = selection.anchorNode || selection.getRangeAt(0).startContainer;
        if (node && node.nodeType === 3) {{
          node = node.parentElement;
        }}
        if (node && node.nodeType === 1) {{
          return node;
        }}
      }}
      if (doc.activeElement && doc.activeElement !== doc.body) {{
        return doc.activeElement;
      }}
      return doc.body;
    }}

    function resolveBorderTarget(element, doc) {{
      if (!element) {{
        return null;
      }}
      if (element.closest) {{
        const cell = element.closest('td, th');
        if (cell) {{
          return cell;
        }}
        const block = element.closest('p, div, section, table, tr, li, h1, h2, h3, h4, h5, h6');
        if (block && block !== doc.body) {{
          return block;
        }}
      }}
      return element === doc.body ? null : element;
    }}

    function buildBorderValue() {{
      const width = (borderWidthInput && borderWidthInput.value) || '1px';
      const style = (borderStyleInput && borderStyleInput.value) || 'solid';
      const color = (borderColorInput && borderColorInput.value) || '#000000';
      return width + ' ' + style + ' ' + color;
    }}

    function applyBorder(side) {{
      const doc = iframe.contentDocument;
      if (!doc) {{
        return;
      }}
      focusEditor();
      const selectedElement = getSelectionElement(doc);
      const target = resolveBorderTarget(selectedElement, doc);
      if (!target) {{
        alert('Place the cursor inside the element you want to style.');
        return;
      }}

      if (side === 'clear') {{
        target.style.border = '';
        target.style.borderTop = '';
        target.style.borderRight = '';
        target.style.borderBottom = '';
        target.style.borderLeft = '';
        return;
      }}

      const borderValue = buildBorderValue();
      if (side === 'all') {{
        target.style.border = borderValue;
        return;
      }}
      if (side === 'top') {{
        target.style.borderTop = borderValue;
        return;
      }}
      if (side === 'right') {{
        target.style.borderRight = borderValue;
        return;
      }}
      if (side === 'bottom') {{
        target.style.borderBottom = borderValue;
        return;
      }}
      if (side === 'left') {{
        target.style.borderLeft = borderValue;
      }}
    }}

    borderButtons.forEach(function (button) {{
      button.addEventListener('click', function () {{
        const side = button.getAttribute('data-border');
        if (!side) {{
          return;
        }}
        applyBorder(side);
      }});
    }});

    undoBtn.addEventListener('click', function () {{
      const doc = iframe.contentDocument;
      if (doc) {{
        focusEditor();
        doc.execCommand('undo', false, null);
      }}
    }});

    redoBtn.addEventListener('click', function () {{
      const doc = iframe.contentDocument;
      if (doc) {{
        focusEditor();
        doc.execCommand('redo', false, null);
      }}
    }});

    resetBtn.addEventListener('click', function () {{
      if (confirm('Reset editor to the revision start state? Unsaved changes will be lost.')) {{
        loadEditorContent();
      }}
    }});

    function submitEditor(saveMode) {{
      const doc = iframe.contentDocument;
      if (!doc || !doc.documentElement) {{
        alert('Editor is not ready yet. Please wait a second and try again.');
        return;
      }}
      normalizeAuditBullets(doc);
      payload.value = '<!doctype html>\\n' + doc.documentElement.outerHTML;
      saveModeInput.value = saveMode || 'session';
      saveBtn.form.submit();
    }}

    saveBtn.addEventListener('click', function () {{
      submitEditor(saveBtn.getAttribute('data-save-mode'));
    }});

    saveNewBtn.addEventListener('click', function () {{
      submitEditor(saveNewBtn.getAttribute('data-save-mode'));
    }});

    if (zoomInput) {{
      zoomInput.addEventListener('change', applyCanvasZoom);
      applyCanvasZoom();
    }}

    loadEditorContent();
  }})();
</script>
"""
        return self._render_editor_shell('Visual Layout Editor', body)

    @http.route('/audit_report/revision/<int:revision_id>/preview', type='http', auth='user')
    def view_revision_preview(self, revision_id, **kwargs):
        revision = self._load_revision(revision_id, include_removed=False)
        if not revision:
            return request.not_found()
        try:
            html_content = revision._build_render_ready_html(refresh_toc=True)
        except ValidationError as e:
            summary, issues = self._extract_error_summary_and_issues(
                e,
                'Unable to render preview for this revision.',
            )
            return self._render_editor_error_response(
                'Preview Unavailable',
                summary,
                issues,
                back_url=f'/web#id={revision.id}&model=audit.report.revision&view_type=form',
                status=400,
            )
        except Exception as e:
            _logger.exception("Saved revision preview generation failed: %s", e)
            summary, issues = self._extract_error_summary_and_issues(
                e,
                'Unable to render preview for this revision.',
            )
            return self._render_editor_error_response(
                'Preview Unavailable',
                summary,
                issues,
                back_url=f'/web#id={revision.id}&model=audit.report.revision&view_type=form',
                status=500,
            )

        return request.make_response(
            html_content,
            headers=[
                ('Content-Type', 'text/html; charset=utf-8'),
                ('Cache-Control', 'no-store, no-cache, must-revalidate, max-age=0'),
                ('Pragma', 'no-cache'),
                ('Expires', '0'),
            ],
        )

    @http.route('/audit_report/revision/<int:revision_id>/pdf', type='http', auth='user')
    def view_revision_pdf(self, revision_id, **kwargs):
        revision = self._load_revision(revision_id, include_removed=False)
        if not revision:
            return request.not_found()

        try:
            pdf_content = revision._get_pdf_content()
        except ValidationError as e:
            summary, issues = self._extract_error_summary_and_issues(
                e,
                'Unable to generate PDF for this revision.',
            )
            return self._render_editor_error_response(
                'Revision PDF Error',
                summary,
                issues,
                back_url=f'/web#id={revision.id}&model=audit.report.revision&view_type=form',
                status=400,
            )
        except Exception as e:
            _logger.exception("Saved revision PDF generation failed: %s", e)
            summary, issues = self._extract_error_summary_and_issues(
                e,
                'Unable to generate PDF for this revision.',
            )
            return self._render_editor_error_response(
                'Revision PDF Error',
                summary,
                issues,
                back_url=f'/web#id={revision.id}&model=audit.report.revision&view_type=form',
                status=500,
            )

        filename = f'Audit_Report_{revision.document_id.id}_v{revision.version_no}.pdf'
        return request.make_response(
            pdf_content,
            headers=[
                ('Content-Type', 'application/pdf'),
                ('Content-Disposition', f'inline; filename=\"{filename}\"'),
                ('Content-Length', str(len(pdf_content))),
                ('Cache-Control', 'no-store, no-cache, must-revalidate, max-age=0'),
                ('Pragma', 'no-cache'),
                ('Expires', '0'),
            ],
        )

    @http.route('/audit_report/revision/<int:revision_id>/tables', type='http', auth='user')
    def list_revision_tables(self, revision_id, **kwargs):
        revision = self._load_revision(revision_id, include_removed=True)
        if not revision:
            return request.not_found()
        payload = {
            'revision_id': revision.id,
            'version_no': revision.version_no,
            'company_id': revision.company_id.id,
            'tables': revision.get_table_index(),
        }
        return request.make_response(
            json.dumps(payload),
            headers=[('Content-Type', 'application/json; charset=utf-8')],
        )

    @http.route('/audit_report/revision/<int:revision_id>/edit/structured', type='http', auth='user')
    def edit_revision_structured(self, revision_id, **kwargs):
        revision = self._load_revision(revision_id, include_removed=False)
        if not revision:
            return request.not_found()
        selected_table_key = kwargs.get('table_key')
        saved_state = str(kwargs.get('saved') or '').strip().lower()
        if saved_state in ('1', 'true', 'yes'):
            saved_state = 'new'
        page = self._render_structured_editor_page(
            revision,
            selected_table_key=selected_table_key,
            saved_state=saved_state,
        )
        return request.make_response(
            page,
            headers=[('Content-Type', 'text/html; charset=utf-8')],
        )

    @http.route(
        '/audit_report/revision/<int:revision_id>/save/structured',
        type='http',
        auth='user',
        methods=['POST'],
    )
    def save_revision_structured(self, revision_id, **kwargs):
        revision = self._load_revision(revision_id, include_removed=False)
        if not revision:
            return request.not_found()

        table_key = kwargs.get('table_key') or ''
        table_action = kwargs.get('table_action') or 'save'
        save_mode = (kwargs.get('save_mode') or 'session').strip().lower()
        force_new_revision = save_mode == 'new'
        target_row = kwargs.get('target_row') or 0
        target_col = kwargs.get('target_col') or 0
        posted_cells = {
            key: value
            for key, value in kwargs.items()
            if key.startswith('cell_')
        }
        back_url = f'/audit_report/revision/{revision.id}/edit/structured'
        if table_key:
            back_url += f'?table_key={quote_plus(table_key)}'

        try:
            html_content = revision.apply_structured_table_changes(
                table_key,
                posted_cells,
                table_action=table_action,
                target_row=target_row,
                target_col=target_col,
            )
            prepared_html = revision.prepare_edited_html_for_storage(html_content)
            target_revision, save_state = self._persist_revision_edit(
                revision,
                prepared_html,
                force_new_revision=force_new_revision,
            )
        except ValidationError as e:
            _logger.warning("Structured table save blocked: %s", e)
            summary, issues = self._extract_error_summary_and_issues(
                e,
                'Could not save structured changes.',
            )
            return self._render_editor_error_response(
                'Structured Save Failed',
                summary,
                issues,
                back_url=back_url,
                status=400,
            )
        except Exception as e:
            _logger.exception("Structured table save failed: %s", e)
            summary, issues = self._extract_error_summary_and_issues(
                e,
                'Unexpected error while saving structured changes.',
            )
            return self._render_editor_error_response(
                'Structured Save Failed',
                summary,
                issues,
                back_url=back_url,
                status=500,
            )

        url = f'/audit_report/revision/{target_revision.id}/edit/structured?saved={save_state}'
        if table_key:
            url += f'&table_key={quote_plus(table_key)}'
        return request.redirect(url)

    @http.route('/audit_report/revision/<int:revision_id>/edit/freeform', type='http', auth='user')
    def edit_revision_freeform(self, revision_id, **kwargs):
        revision = self._load_revision(revision_id, include_removed=False)
        if not revision:
            return request.not_found()

        saved_state = str(kwargs.get('saved') or '').strip().lower()
        if saved_state in ('1', 'true', 'yes'):
            saved_state = 'new'
        page = self._render_freeform_editor_page(revision, saved_state=saved_state)
        return request.make_response(
            page,
            headers=[('Content-Type', 'text/html; charset=utf-8')],
        )

    @http.route(
        '/audit_report/revision/<int:revision_id>/save/freeform',
        type='http',
        auth='user',
        methods=['POST'],
    )
    def save_revision_freeform(self, revision_id, **kwargs):
        revision = self._load_revision(revision_id, include_removed=False)
        if not revision:
            return request.not_found()

        html_content = kwargs.get('html_content') or ''
        save_mode = (kwargs.get('save_mode') or 'session').strip().lower()
        force_new_revision = save_mode == 'new'
        back_url = f'/audit_report/revision/{revision.id}/edit/freeform'
        if not html_content.strip():
            return self._render_editor_error_response(
                'Freeform Save Failed',
                'Freeform HTML cannot be empty.',
                back_url=back_url,
                status=400,
            )
        try:
            prepared_html = revision.prepare_edited_html_for_storage(html_content)
            target_revision, save_state = self._persist_revision_edit(
                revision,
                prepared_html,
                force_new_revision=force_new_revision,
            )
        except ValidationError as e:
            _logger.warning("Freeform save blocked: %s", e)
            summary, issues = self._extract_error_summary_and_issues(
                e,
                'Could not save freeform changes.',
            )
            return self._render_editor_error_response(
                'Freeform Save Failed',
                summary,
                issues,
                back_url=back_url,
                status=400,
            )
        except Exception as e:
            _logger.exception("Freeform save failed: %s", e)
            summary, issues = self._extract_error_summary_and_issues(
                e,
                'Unexpected error while saving freeform changes.',
            )
            return self._render_editor_error_response(
                'Freeform Save Failed',
                summary,
                issues,
                back_url=back_url,
                status=500,
            )

        return request.redirect(f'/audit_report/revision/{target_revision.id}/edit/freeform?saved={save_state}')

    @http.route(
        '/audit_report/revision/<int:revision_id>/remove',
        type='http',
        auth='user',
        methods=['POST'],
    )
    def remove_revision(self, revision_id, **kwargs):
        revision = self._load_revision(revision_id, include_removed=True)
        if not revision:
            return request.not_found()
        revision.action_soft_remove()
        next_url = kwargs.get('next') or (
            f'/web#id={revision.document_id.id}&model=audit.report.document&view_type=form'
        )
        return request.redirect(next_url)

    @http.route(
        '/audit_report/revision/<int:revision_id>/restore',
        type='http',
        auth='user',
        methods=['POST'],
    )
    def restore_revision(self, revision_id, **kwargs):
        revision = self._load_revision(revision_id, include_removed=True)
        if not revision:
            return request.not_found()
        revision.action_restore()
        next_url = kwargs.get('next') or (
            f'/web#id={revision.id}&model=audit.report.revision&view_type=form'
        )
        return request.redirect(next_url)
