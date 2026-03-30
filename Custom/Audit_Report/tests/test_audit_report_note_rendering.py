import json
from types import SimpleNamespace
from unittest.mock import patch

from odoo import fields
from odoo.tests import TransactionCase, tagged

from odoo.addons.Audit_Report.controllers.main import AuditReportController


@tagged('post_install', '-at_install')
class TestAuditReportNoteRendering(TransactionCase):
    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        cls.company = cls.env.company
        cls.AuditReportModel = type(cls.env['audit.report'])
        cls.controller = AuditReportController()

    def _wizard_values(self, **overrides):
        values = {
            'company_id': self.company.id,
            'date_start': fields.Date.to_date('2024-01-01'),
            'date_end': fields.Date.to_date('2024-12-31'),
            'balance_sheet_date_mode': 'end_only',
            'prior_year_mode': 'auto',
            'prior_balance_sheet_date_mode': 'end_only',
            'report_type': 'period',
            'audit_period_category': 'normal_1y',
            'auditor_type': 'default',
            'signature_date_mode': 'today',
            'use_previous_settings': False,
        }
        values.update(overrides)
        return values

    def _create_wizard(self, **overrides):
        return self.env['audit.report'].create(self._wizard_values(**overrides))

    def _create_document_with_revision(self, snapshot_json):
        document = self.env['audit.report.document'].create({
            'name': 'Note Render Test Report',
            'company_id': self.company.id,
            'date_start': fields.Date.to_date('2024-01-01'),
            'date_end': fields.Date.to_date('2024-12-31'),
            'report_type': 'period',
            'audit_period_category': 'normal_1y',
            'source_wizard_json': snapshot_json,
        })
        revision = document.create_revision_from_html('<html><body><p>Base</p></body></html>')
        return revision

    def _render_notes_only_template(
        self,
        template_name,
        *,
        show_prior_year=False,
        ignore_notes_last_page_margins=False,
        note_sections=None,
        note_numbers=None,
        ppe_note_number=False,
        ppe_note_columns=None,
        ppe_note_schedules=None,
        show_restated_headers=False,
        show_related_parties_note=False,
        related_party_rows=None,
    ):
        wizard = self._create_wizard(
            audit_period_category='normal_2y' if show_prior_year else 'normal_1y',
            ignore_notes_last_page_margins=ignore_notes_last_page_margins,
        )
        lines = [
            {
                'code': f'5102010{idx}',
                'name': f'Line {idx}',
                'current': float(idx * 10),
                'prev': float(idx * 5),
            }
            for idx in range(1, 8)
        ]
        note = {
            'number': 5,
            'label': 'Operating expenses',
            'lines': lines,
            'line_segments': wizard._build_generic_note_render_segments(lines),
            'total_current': sum(line['current'] for line in lines),
            'total_prev': sum(line['prev'] for line in lines),
            'preserve_sign': False,
        }
        rendered_note_sections = note_sections if note_sections is not None else [note]
        rendered_note_numbers = note_numbers if note_numbers is not None else {'pl_opex': 5}
        env = self.controller._get_template_env(self.controller._templates_path())
        template = env.get_template(template_name)
        return template.render(
            sections_to_render=['notes_to_financial_statements'],
            note_numbers=rendered_note_numbers,
            note_sections=rendered_note_sections,
            ppe_note_number=ppe_note_number,
            ppe_note_columns=ppe_note_columns or [],
            ppe_note_schedules=ppe_note_schedules or [],
            show_prior_year=show_prior_year,
            show_restated_headers=show_restated_headers,
            gap_notes_to_financial_statements=True,
            ignore_notes_last_page_margins=ignore_notes_last_page_margins,
            show_related_parties_note=show_related_parties_note,
            show_shareholder_note=True,
            show_share_capital_conversion_note=False,
            show_share_capital_transfer_note=False,
            signature_break_lines_notes=2,
            signature_names=['Director One'],
            company_name='Demo Co',
            generated_date_display='01 January 2025',
            report_period_word='period',
            comparative_period_word='period',
            share_capital_paid_status='paid',
            share_rows=[],
            authorized_share_capital=0.0,
            total_shares_count=0.0,
            share_value_default=0.0,
            related_party_rows=related_party_rows or [],
            is_dmcc_company_freezone=False,
            date_end=fields.Date.to_date('2024-12-31'),
        )

    def _render_balance_sheet_only_template(self, template_name, *, show_restated_headers=False):
        env = self.controller._get_template_env(self.controller._templates_path())
        template = env.get_template(template_name)
        return template.render(
            sections_to_render=['balance_sheet_page'],
            company_name='Demo Co',
            date_end=fields.Date.to_date('2024-12-31'),
            current_group_totals={},
            prev_group_totals={},
            main_head_labels={},
            note_numbers={},
            total_assets=0.0,
            total_liabilities=0.0,
            total_equity=0.0,
            total_of_equity_and_liabilities=0.0,
            prev_total_assets=0.0,
            prev_total_liabilities=0.0,
            prev_total_equity=0.0,
            prev_total_of_equity_and_liabilities=0.0,
            tb_diff_current=0.0,
            tb_diff_prior=0.0,
            tb_warning_current=False,
            tb_warning_prior=False,
            show_prior_year=True,
            show_restated_headers=show_restated_headers,
            is_dormant_period=False,
            is_cessation_period=False,
        )

    def _render_auditor_report_template(self, template_name, *, note_ref=5):
        env = self.controller._get_template_env(self.controller._templates_path())
        template = env.get_template(template_name)
        return template.render(
            sections_to_render=['independent_auditor_report'],
            gap_independent_auditor_report=True,
            report_ended_label='Period Ended',
            date_end=fields.Date.to_date('2024-12-31'),
            company_name='Demo Co',
            report_period_word='period',
            generated_date_display='01 January 2025',
            freezone_selection='default',
            show_emphasis_of_matter=True,
            emphasis_note_items=[{
                'note_ref': note_ref,
                'matter_text': 'the correction of a prior period error',
            }],
            is_default=True,
            is_ifza=False,
        )

    def _render_entity_information_template(self, template_name, owner_display_names):
        env = self.controller._get_template_env(self.controller._templates_path())
        template = env.get_template(template_name)
        return template.render(
            sections_to_render=['entity_information'],
            freezone_selection='default',
            freezone='IFZA',
            signature_names=[],
            report_ended_label='Period Ended',
            report_period_end='31 December 2025',
            owner='',
            owner_display_names=owner_display_names,
            company=SimpleNamespace(street='DSO-IFZA, IFZA Properties', city='Dubai'),
            license='53589',
        )

    @staticmethod
    def _normalize_html_text(html):
        return ' '.join((html or '').split())

    def _render_report_of_directors_html(self, wizard, *, report_data=None):
        return self.controller._render_report_html(
            wizard,
            sections_to_render=['report_of_directors'],
            toc_entries=[],
            report_data=report_data,
            css_content='',
        )

    def _render_report_sections_html(self, wizard, sections_to_render, *, report_data=None):
        return self.controller._render_report_html(
            wizard,
            sections_to_render=sections_to_render,
            toc_entries=[],
            report_data=report_data,
            css_content='',
        )

    def test_collect_note_lines_merges_vat_receivable_and_payable_prefixes(self):
        wizard = self._create_wizard()
        current_map = {
            '1203': {
                '12030101': {'name': 'Advance', 'balance': 15.0},
                '12030401': {'name': 'Input VAT - Main', 'balance': 10.0},
                '12030402': {'name': 'Input VAT - Reverse Charge', 'balance': 5.0},
            },
            '2202': {
                '22020101': {'name': 'Trade Payable', 'balance': 22.0},
            },
            '2203': {
                '22030101': {'name': 'Other Payable', 'balance': 12.0},
                '22030201': {'name': 'Output VAT - Standard Rated', 'balance': 7.0},
                '22030202': {'name': 'Output VAT - Reverse Charge', 'balance': 3.0},
                '22030301': {'name': 'Audit Fee Accrual', 'balance': 5.0},
            },
        }

        receivable_lines = wizard._collect_note_lines(['1203'], current_map, {})
        payable_lines = wizard._collect_note_lines(['2201', '2202', '2203', '2204'], current_map, {})

        vat_receivable = next(line for line in receivable_lines if line['name'] == 'VAT recoverable')
        vat_payable = next(line for line in payable_lines if line['name'] == 'VAT payable')
        audit_fee_accrual = next(
            line for line in payable_lines
            if line['name'] == 'Audit and Accounting fee accrual'
        )

        self.assertEqual(vat_receivable['current'], 15.0)
        self.assertEqual(vat_payable['current'], 10.0)
        self.assertEqual(audit_fee_accrual['current'], 5.0)
        self.assertFalse(any('Input VAT' in line['name'] for line in receivable_lines))
        self.assertFalse(any('Output VAT' in line['name'] for line in payable_lines))
        self.assertFalse(any(line['name'] == 'Audit Fee Accrual' for line in payable_lines))

    def test_operating_expense_normalization_merges_requested_labels(self):
        wizard = self._create_wizard()
        normalized = wizard._normalize_operating_expense_note_lines([
            {'code': '51020101', 'name': 'Subcontractor Services', 'current': 10.0, 'prev': 1.0},
            {'code': '51020102', 'name': 'Entertainment Expense', 'current': 8.0, 'prev': 2.0},
            {'code': '51090101', 'name': 'Audit and Accounting Fee', 'current': 11.0, 'prev': 5.0},
            {'code': '51090102', 'name': 'Other Accountant Fee', 'current': 9.0, 'prev': 4.0},
            {'code': '51090103', 'name': 'Audit Fee Accrual', 'current': 6.0, 'prev': 3.0},
            {'code': '51130101', 'name': 'Business Travel', 'current': 7.0, 'prev': 3.0},
            {'code': '51130201', 'name': 'Business Accommodation', 'current': 5.0, 'prev': 4.0},
            {'code': '51250101', 'name': 'Business Insurance Expense', 'current': 6.0, 'prev': 1.0},
            {'code': '51030101', 'name': 'IT Expenses', 'current': 3.0, 'prev': 2.0},
            {'code': '51030102', 'name': 'Software Subscriptions', 'current': 4.0, 'prev': 1.0},
        ])
        names = {line['name']: line for line in normalized}

        self.assertIn('Subcontractor', names)
        self.assertIn('Entertainment', names)
        self.assertIn('Audit and accounting fee', names)
        self.assertIn('Audit and Accounting fee accrual', names)
        self.assertIn('Travelling and accommodation', names)
        self.assertIn('Insurance expense', names)
        self.assertIn('IT expenses', names)
        self.assertEqual(names['Audit and accounting fee']['current'], 20.0)
        self.assertEqual(names['Audit and accounting fee']['prev'], 9.0)
        self.assertEqual(names['Audit and Accounting fee accrual']['current'], 6.0)
        self.assertEqual(names['Travelling and accommodation']['current'], 12.0)
        self.assertEqual(names['IT expenses']['current'], 7.0)
        self.assertNotIn('Software Subscriptions', names)
        self.assertNotIn('Other Accountant Fee', names)
        self.assertNotIn('Audit Fee Accrual', names)

    def test_cost_of_revenue_normalization_renames_requested_labels(self):
        wizard = self._create_wizard()

        normalized = wizard._normalize_cost_of_revenue_note_lines([
            {'code': '51010101', 'name': 'Stock Purchase', 'current': 40.0, 'prev': 10.0},
            {'code': '51010102', 'name': 'Subcontractor Services', 'current': 4.0, 'prev': 1.0},
        ])

        self.assertEqual(normalized[0]['name'], 'Purchases')
        self.assertEqual(normalized[1]['name'], 'Subcontractor')

    def test_canonical_note_name_renames_staff_salary_variants(self):
        wizard = self._create_wizard()

        self.assertEqual(
            wizard._canonical_note_line_display_name('Staff Salary'),
            'Salaries and wages',
        )
        self.assertEqual(
            wizard._canonical_note_line_display_name('Coaching Staff Salaries'),
            'Salaries and wages',
        )

    def test_full_year_body_period_word_uses_date_span_while_header_stays_report_type(self):
        wizard = self._create_wizard(
            report_type='period',
            date_start=fields.Date.to_date('2024-01-01'),
            date_end=fields.Date.to_date('2024-12-31'),
        )

        report_data = wizard._get_report_data()
        rendered = self._normalize_html_text(
            self._render_report_of_directors_html(wizard, report_data=report_data)
        )

        self.assertEqual(report_data['report_period_word'], 'year')
        self.assertIn('For the Period Ended 31 December 2024 (in AED)', rendered)
        self.assertIn('financial statements for the year ended 31 December 2024.', rendered)

    def test_partial_year_body_period_word_uses_date_span_while_header_stays_report_type(self):
        wizard = self._create_wizard(
            report_type='year',
            date_start=fields.Date.to_date('2024-04-01'),
            date_end=fields.Date.to_date('2024-06-30'),
        )

        report_data = wizard._get_report_data()
        rendered = self._normalize_html_text(
            self._render_report_of_directors_html(wizard, report_data=report_data)
        )
        soce_profit_row = next(
            row for row in report_data['soce_rows']
            if (row.get('label') or '').startswith('Total comprehensive')
        )

        self.assertEqual(report_data['report_period_word'], 'period')
        self.assertEqual(soce_profit_row.get('period_word'), 'period')
        self.assertIn('For the Year Ended 30 June 2024 (in AED)', rendered)
        self.assertIn('financial statements for the period ended 30 June 2024.', rendered)

    def test_generic_note_render_segments_keep_two_rows_at_start_and_end(self):
        wizard = self._create_wizard()
        lines = [
            {'code': f'5102010{idx}', 'name': f'Line {idx}', 'current': float(idx), 'prev': 0.0}
            for idx in range(1, 8)
        ]

        segments = wizard._build_generic_note_render_segments(lines)

        self.assertEqual([len(segment['lines']) for segment in segments], [2, 2, 3])
        self.assertTrue(segments[0]['show_title'])
        self.assertFalse(segments[0]['show_total'])
        self.assertFalse(segments[1]['show_title'])
        self.assertFalse(segments[1]['show_total'])
        self.assertTrue(segments[-1]['show_total'])

    def test_revision_snapshot_restores_ignore_notes_last_page_margins(self):
        wizard = self._create_wizard(ignore_notes_last_page_margins=True)
        snapshot_json = wizard._get_wizard_snapshot_json()
        revision = self._create_document_with_revision(snapshot_json)

        with patch.object(self.AuditReportModel, '_load_tb_override_lines', lambda *args, **kwargs: None), \
                patch.object(self.AuditReportModel, '_apply_tb_overrides_from_serialized_payload', lambda *args, **kwargs: None), \
                patch.object(self.AuditReportModel, '_sync_tb_overrides_json', lambda *args, **kwargs: ''), \
                patch.object(self.AuditReportModel, '_apply_lor_extra_items_from_serialized_payload', lambda *args, **kwargs: None), \
                patch.object(self.AuditReportModel, '_get_report_data', lambda *args, **kwargs: {}):
            restored_wizard = revision._build_audit_report_wizard_from_snapshot()

        self.assertTrue(json.loads(snapshot_json)['ignore_notes_last_page_margins'])
        self.assertTrue(restored_wizard.ignore_notes_last_page_margins)

    def test_revision_snapshot_restores_business_activity_include_providing(self):
        wizard = self._create_wizard(business_activity_include_providing=False)
        snapshot_json = wizard._get_wizard_snapshot_json()
        revision = self._create_document_with_revision(snapshot_json)

        with patch.object(self.AuditReportModel, '_load_tb_override_lines', lambda *args, **kwargs: None), \
                patch.object(self.AuditReportModel, '_apply_tb_overrides_from_serialized_payload', lambda *args, **kwargs: None), \
                patch.object(self.AuditReportModel, '_sync_tb_overrides_json', lambda *args, **kwargs: ''), \
                patch.object(self.AuditReportModel, '_apply_lor_extra_items_from_serialized_payload', lambda *args, **kwargs: None), \
                patch.object(self.AuditReportModel, '_get_report_data', lambda *args, **kwargs: {}):
            restored_wizard = revision._build_audit_report_wizard_from_snapshot()

        self.assertFalse(json.loads(snapshot_json)['business_activity_include_providing'])
        self.assertFalse(restored_wizard.business_activity_include_providing)

    def test_business_activity_words_can_be_toggled_across_report_and_notes(self):
        self.company.trade_license_activities = 'Consulting'
        wizard = self._create_wizard(
            business_activity_include_providing=False,
            business_activity_include_services=False,
        )

        report_data = wizard._get_report_data()
        rendered = self._normalize_html_text(
            self._render_report_sections_html(
                wizard,
                ['report_of_directors', 'notes_to_financial_statements'],
                report_data=report_data,
            )
        )

        self.assertEqual(rendered.count('The business activity of the Entity is consulting.'), 2)
        self.assertNotIn('providing consulting', rendered)
        self.assertNotIn('consulting services', rendered)

    def test_dmcc_summary_sheet_renders_dynamic_header_and_portal_account(self):
        self.company.free_zone = 'Dubai Multi Commodities Centre Free Zone'
        wizard = self._create_wizard(report_type='year', portal_account_no='368661')

        rendered = self._normalize_html_text(
            self._render_report_sections_html(
                wizard,
                ['dmcc_sheet'],
                report_data=wizard._get_report_data(),
            )
        )

        self.assertIn('DMCC Summary Sheet', rendered)
        self.assertIn('For the Year Ended 31 December 2024 (in AED)', rendered)
        self.assertIn('Portal Account', rendered)
        self.assertIn('368661', rendered)

    def test_soce_first_balance_label_date_is_manual_only(self):
        wizard = self._create_wizard(audit_period_category='normal_2y')

        report_data = wizard._get_report_data()
        first_balance_row = report_data['soce_rows'][0]

        self.assertEqual(first_balance_row['label'], 'Balance as at ')
        self.assertNotIn('01 January 2023', first_balance_row['label'])
        self.assertIn('set it manually', wizard.soce_warning_message)

    def test_notes_template_renders_segmented_tbodies_and_bottom_margin_page_class(self):
        html_default = self._render_notes_only_template(
            'audit_report_template.html',
            show_prior_year=False,
            ignore_notes_last_page_margins=False,
        )
        html_compact = self._render_notes_only_template(
            'audit_report_template.html',
            show_prior_year=False,
            ignore_notes_last_page_margins=True,
        )

        self.assertEqual(html_default.count('note-block note-block-segment'), 3)
        self.assertEqual(html_compact.count('note-block note-block-segment'), 3)
        self.assertEqual(html_default.count('<br>'), 2)
        self.assertEqual(html_compact.count('<br>'), 2)
        self.assertIn('notes-ignore-bottom-margin', html_compact)
        self.assertNotIn('notes-ignore-bottom-margin', html_default)

    def test_notes_template_2y_renders_segmented_tbodies(self):
        html = self._render_notes_only_template(
            'audit_report_template_2y.html',
            show_prior_year=True,
            ignore_notes_last_page_margins=False,
        )

        self.assertEqual(html.count('note-block note-block-segment'), 3)

    def test_ppe_note_renders_explicit_equal_width_value_columns(self):
        ppe_note_columns = [
            {'code': 'furniture', 'label': 'Furniture and fixtures'},
            {'code': 'it', 'label': 'IT equipments'},
            {'code': 'office', 'label': 'Office equipments'},
            {'code': 'total', 'label': 'Total'},
        ]
        ppe_note_schedules = [{
            'rows': [
                {
                    'row_type': 'section',
                    'label': 'Cost',
                    'values': [None, None, None, None],
                },
                {
                    'row_type': 'line',
                    'label': 'As at 01 January 2025',
                    'values': [18158.0, 1599.0, 0.0, 19757.0],
                },
            ],
        }]
        html = self._render_notes_only_template(
            'audit_report_template.html',
            note_sections=[{
                'number': 5,
                'label': 'Property plant and equipment',
                'lines': [],
                'line_segments': [],
                'total_current': 0.0,
                'total_prev': 0.0,
                'preserve_sign': False,
            }],
            note_numbers={'1101': 5},
            ppe_note_number=5,
            ppe_note_columns=ppe_note_columns,
            ppe_note_schedules=ppe_note_schedules,
        )

        self.assertIn('<colgroup>', html)
        self.assertIn('class="ppe-note-col-label"', html)
        self.assertEqual(html.count('class="ppe-note-col-value"'), len(ppe_note_columns))
        self.assertIn('--ppe-value-column-count: 4;', html)

    def test_previous_settings_round_trip_preserves_correction_error_payload(self):
        wizard = self._create_wizard(
            emphasis_correction_error=True,
            correction_error_note_body='Paragraph one\nParagraph two',
        )
        wizard.correction_error_line_ids = wizard._correction_error_rows_to_commands(
            [
                {'sequence': 10, 'row_type': 'section', 'description': 'Effect on statement of financial position'},
                {'sequence': 20, 'row_type': 'line', 'description': 'Cash and bank balance', 'amount_as_reported': 10.0, 'amount_as_restated': 15.0, 'amount_restatement': 5.0},
            ],
            clear_existing=True,
        )

        wizard._store_previous_settings()

        restored = self._create_wizard(use_previous_settings=True)
        restored._onchange_use_previous_settings()

        self.assertTrue(restored.emphasis_correction_error)
        self.assertEqual(restored.correction_error_note_body, 'Paragraph one\nParagraph two')
        self.assertEqual(len(restored.correction_error_line_ids), 2)
        self.assertEqual(restored.correction_error_line_ids[0].row_type, 'section')
        self.assertEqual(restored.correction_error_line_ids[1].description, 'Cash and bank balance')

    def test_notes_template_renders_correction_error_note_before_regular_notes(self):
        wizard = self._create_wizard(audit_period_category='normal_2y')
        regular_lines = [
            {'code': '51020101', 'name': 'Administrative expense', 'current': 40.0, 'prev': 30.0},
            {'code': '51020102', 'name': 'Professional fee', 'current': 20.0, 'prev': 10.0},
        ]
        correction_note = {
            'key': 'correction_error',
            'number': 5,
            'label': 'Correction of Error',
            'paragraphs': ['Prior period balances were corrected.', 'Comparatives were re-stated accordingly.'],
            'correction_header_date_display': '31 December 2023',
            'correction_rows': [
                {'sequence': 10, 'row_type': 'section', 'description': 'Effect on statement of financial position'},
                {'sequence': 20, 'row_type': 'subheading', 'description': 'Current assets'},
                {'sequence': 30, 'row_type': 'line', 'description': 'Cash and bank balance', 'amount_as_reported': 354163.0, 'amount_as_restated': 765463.0, 'amount_restatement': 411300.0},
                {'sequence': 40, 'row_type': 'text', 'description': 'There is no effect on statement of comprehensive income'},
            ],
        }
        regular_note = {
            'number': 6,
            'label': 'Operating expenses',
            'lines': regular_lines,
            'line_segments': wizard._build_generic_note_render_segments(regular_lines),
            'total_current': 60.0,
            'total_prev': 40.0,
            'preserve_sign': False,
        }

        html = self._render_notes_only_template(
            'audit_report_template_2y.html',
            show_prior_year=True,
            note_sections=[correction_note, regular_note],
            note_numbers={'correction_error': 5, 'pl_opex': 6},
            show_restated_headers=True,
        )

        self.assertIn('Correction of Error', html)
        self.assertIn('As re-stated', html)
        self.assertIn('Effect on statement of financial position', html)
        self.assertIn('correction-error-date-heading', html)
        self.assertIn('correction-error-currency-head', html)
        self.assertLess(html.index('Effect on statement of financial position'), html.index('Current assets'))
        self.assertLess(html.index('Current assets'), html.index('Cash and bank balance'))
        self.assertLess(html.index('Correction of Error'), html.index('Operating expenses'))
        self.assertIn('Re-stated', html)

    def test_balance_sheet_template_renders_restated_header_only_when_enabled(self):
        html_default = self._render_balance_sheet_only_template(
            'audit_report_template_2y.html',
            show_restated_headers=False,
        )
        html_restated = self._render_balance_sheet_only_template(
            'audit_report_template_2y.html',
            show_restated_headers=True,
        )

        self.assertNotIn('Re-stated', html_default)
        self.assertIn('Re-stated', html_restated)

    def test_auditor_report_template_places_emphasis_after_basis_for_opinion(self):
        html = self._render_auditor_report_template(
            'audit_report_template_2y.html',
            note_ref=5,
        )

        self.assertIn('Basis for Opinion:', html)
        self.assertIn('Emphasis of Matter:', html)
        self.assertIn('We draw attention to Note (5)', html)
        self.assertLess(html.index('Basis for Opinion:'), html.index('Emphasis of Matter:'))

    def test_entity_information_adds_spacing_class_for_more_than_four_owners(self):
        html_without_extra_spacing = self._render_entity_information_template(
            'audit_report_template.html',
            ['Owner 1', 'Owner 2', 'Owner 3', 'Owner 4'],
        )
        html_with_extra_spacing = self._render_entity_information_template(
            'audit_report_template.html',
            ['Owner 1', 'Owner 2', 'Owner 3', 'Owner 4', 'Owner 5'],
        )

        self.assertNotIn('entity-owner-row entity-owner-row-spacious', html_without_extra_spacing)
        self.assertIn('entity-owner-row entity-owner-row-spacious', html_with_extra_spacing)
