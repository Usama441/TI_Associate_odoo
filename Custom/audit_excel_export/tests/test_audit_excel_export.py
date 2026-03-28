import base64
import datetime
import io
import unittest
from unittest.mock import patch

try:
    from openpyxl import Workbook, load_workbook
except ImportError:
    Workbook = load_workbook = None

from odoo import Command, fields
from odoo.addons.audit_excel_export.utils import clean_bank_narration
from odoo.addons.account.tests.common import AccountTestInvoicingCommon
from odoo.exceptions import ValidationError
from odoo.tests.common import tagged


@tagged('post_install', '-at_install')
@unittest.skipUnless(load_workbook, 'openpyxl is required for XLSX assertions')
class TestAuditExcelExport(AccountTestInvoicingCommon):

    @classmethod
    def setUpClass(cls):
        super().setUpClass()

        cls.export_date_from = fields.Date.from_string('2025-01-01')
        cls.export_date_to = fields.Date.from_string('2025-12-31')
        cls.aged_as_of_date = fields.Date.from_string('2025-12-31')

        cls.customer_partner = cls.partner_a
        cls.vendor_partner = cls.partner_b

        cls.customer_invoice_posted = cls.init_invoice(
            'out_invoice',
            partner=cls.customer_partner,
            invoice_date=fields.Date.from_string('2025-01-10'),
            post=True,
            amounts=[100.0],
        )
        cls.vendor_bill_posted = cls.init_invoice(
            'in_invoice',
            partner=cls.vendor_partner,
            invoice_date=fields.Date.from_string('2025-01-14'),
            post=True,
            amounts=[80.0],
        )

    def _create_wizard(self, **overrides):
        vals = {
            'company_ids': [Command.set(self.env.company.ids)],
            'date_from': self.export_date_from,
            'date_to': self.export_date_to,
            'aged_as_of_date': self.aged_as_of_date,
            'export_sheet_key': 'all',
            'include_draft_entries': True,
            'unfold_all': True,
            'hide_zero_lines': False,
            'aging_based_on': 'base_on_maturity_date',
            'aging_interval': 30,
            'show_currency': True,
            'show_account': True,
            'include_dynamic_columns': True,
            'invoice_bill_scope': 'all_states',
            'include_refunds': True,
            'partner_ids': [Command.set([self.customer_partner.id, self.vendor_partner.id])],
        }
        vals.update(overrides)
        return self.env['audit.excel.export.wizard'].create(vals)

    def _export_workbook(self, wizard):
        action = wizard.action_export_xlsx()
        self.assertEqual(action.get('type'), 'ir.actions.act_url')
        self.assertTrue(wizard.file_data)
        binary = base64.b64decode(wizard.file_data)
        self.assertGreater(len(binary), 0)
        wb = load_workbook(io.BytesIO(binary), data_only=False)
        return action, wb

    def _sheet_values(self, ws):
        values = []
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    values.append(str(cell.value))
        return values

    def _find_header_position(self, ws, header_name, *, max_scan_rows=80):
        for row_idx in range(1, min(ws.max_row, max_scan_rows) + 1):
            for col_idx in range(1, ws.max_column + 1):
                if ws.cell(row=row_idx, column=col_idx).value == header_name:
                    return row_idx, col_idx
        return None, None

    def _normalize_label(self, value):
        return ''.join(ch for ch in str(value or '').strip().lower() if ch.isalnum())

    def _find_row_by_label(self, ws, label, *, start_row=1, occurrence=1):
        expected = self._normalize_label(label)
        seen = 0
        for row_idx in range(start_row, ws.max_row + 1):
            if self._normalize_label(ws.cell(row=row_idx, column=1).value) != expected:
                continue
            seen += 1
            if seen == occurrence:
                return row_idx
        return None

    def _find_row_by_keys(self, ws, keys, *, start_row=1):
        normalized = {self._normalize_label(key) for key in keys}
        for row_idx in range(start_row, ws.max_row + 1):
            if self._normalize_label(ws.cell(row=row_idx, column=1).value) in normalized:
                return row_idx
        return None

    def _coerce_openpyxl_date(self, value):
        if isinstance(value, datetime.datetime):
            return value.date()
        return value

    def _static_numeric_like_values(self, ws):
        values = []
        for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 250), min_col=1, max_col=min(ws.max_column, 8)):
            for cell in row:
                value = cell.value
                if value is None:
                    continue
                if isinstance(value, str) and value.startswith('='):
                    continue
                if isinstance(value, (int, float, bool, datetime.date, datetime.datetime)):
                    values.append((cell.coordinate, value))
        return values

    def test_01_export_download_and_sheet_order(self):
        wizard = self._create_wizard()
        action, wb = self._export_workbook(wizard)

        self.assertIn('download=true', action.get('url', ''))
        self.assertTrue(wizard.file_name.endswith('.xlsx'))
        self.assertEqual(
            wb.sheetnames,
            [
                'Client Details',
                'General Ledger',
                'Trial Balance',
                'SOFP',
                'SOCI',
                'SOCE',
                'SOCF',
                'Prepayment',
                'Customer Invoices',
                'Aged Receivables',
                'Vendor Bills',
                'Aged Payables',
                'VAT Control',
                'Accruals',
                'Summary Sheet',
                'Share Capital',
            ],
        )

    def test_02_template_formulas_and_formatting_survive(self):
        wizard = self._create_wizard()
        _action, wb = self._export_workbook(wizard)

        self.assertTrue(str(wb['Summary Sheet']['A1'].value).startswith('='))
        self.assertTrue(str(wb['Summary Sheet']['B7'].value).startswith('='))
        self.assertTrue(str(wb['SOFP']['A1'].value).startswith('='))
        self.assertTrue(str(wb['VAT Control']['B18'].value).startswith('='))

        sofp = wb['SOFP']
        self.assertIn('A2:E2', {str(rng) for rng in sofp.merged_cells.ranges})
        self.assertTrue(sofp['A2'].font.bold)
        self.assertEqual(sofp.page_setup.orientation, 'landscape')
        self.assertEqual(sofp.column_dimensions['B'].width, 18.0)
        self.assertEqual(sofp.row_dimensions[2].height, 28.0)

    def test_03_first_five_sheets_have_data(self):
        wizard = self._create_wizard()
        _action, wb = self._export_workbook(wizard)

        gl_values = self._sheet_values(wb['General Ledger'])
        self.assertIn('Code', gl_values)
        self.assertIn('Account Name', gl_values)
        self.assertIn('Debit', gl_values)
        self.assertIn('Credit', gl_values)
        self.assertIn('Balance', gl_values)
        self.assertGreater(len(gl_values), 20)

        customer_ws = wb['Customer Invoices']
        customer_values = self._sheet_values(customer_ws)
        expected_summary_headers = [
            'Invoice No',
            'Invoice Date',
            'Accounting Date',
            'Currency',
            'Customer',
            'Due Date',
            'Conversion Rate',
        ]
        expected_line_headers = [
            'Label',
            'Account',
            'Quantity',
            'Price',
            'Taxes',
            'VAT Amount',
            'Amount',
            'Currency',
        ]
        for header in expected_summary_headers:
            self.assertIn(header, customer_values)
        for header in expected_line_headers:
            self.assertIn(header, customer_values)
        self.assertGreater(len(customer_values), 20)
        self.assertIn(self.customer_invoice_posted.name, customer_values)

        vendor_ws = wb['Vendor Bills']
        vendor_values = self._sheet_values(vendor_ws)
        for header in ('Invoice No', 'Invoice Date', 'Accounting Date', 'Currency', 'Vendor', 'Due Date', 'Conversion Rate'):
            self.assertIn(header, vendor_values)
        for header in expected_line_headers:
            self.assertIn(header, vendor_values)
        self.assertGreater(len(vendor_values), 20)
        self.assertIn(self.vendor_bill_posted.name, vendor_values)

        for sheet_name, ws in (('Customer Invoices', customer_ws), ('Vendor Bills', vendor_ws)):
            header_row, line_description_col = self._find_header_position(ws, 'Label')
            self.assertIsNotNone(header_row)
            has_line_rows = any(
                ws.cell(row=row_idx, column=line_description_col).value not in (None, '')
                for row_idx in range(header_row + 1, min(ws.max_row, header_row + 250) + 1)
            )
            self.assertTrue(has_line_rows, f'Expected invoice/bill line rows in {sheet_name}')

            for numeric_header in ('Quantity', 'Price', 'VAT Amount', 'Amount'):
                _hdr_row, numeric_col = self._find_header_position(ws, numeric_header)
                self.assertIsNotNone(numeric_col)
                has_numeric_value = any(
                    isinstance(ws.cell(row=row_idx, column=numeric_col).value, (int, float))
                    for row_idx in range(header_row + 1, min(ws.max_row, header_row + 250) + 1)
                )
                self.assertTrue(has_numeric_value, f'Expected numeric values under {numeric_header} in {sheet_name}')

            _hdr_row, conversion_col = self._find_header_position(ws, 'Conversion Rate')
            self.assertIsNotNone(conversion_col)
            has_conversion_value = any(
                isinstance(ws.cell(row=row_idx, column=conversion_col).value, (int, float))
                for row_idx in range(_hdr_row + 1, min(ws.max_row, _hdr_row + 250) + 1)
            )
            self.assertTrue(has_conversion_value, f'Expected conversion rate values in {sheet_name}')

        for sheet_name in ('Aged Receivables', 'Aged Payables'):
            ws = wb[sheet_name]
            values = self._sheet_values(ws)
            self.assertGreater(ws.max_row, 5)
            self.assertGreater(ws.max_column, 5)
            self.assertGreater(len(values), 20)

    def test_04_trial_balance_is_populated(self):
        wizard = self._create_wizard()
        _action, wb = self._export_workbook(wizard)

        ws = wb['Trial Balance']
        values = self._sheet_values(ws)
        self.assertIn('Code', values)
        self.assertIn('Account Name', values)
        self.assertIn('Debit', values)
        self.assertIn('Credit', values)
        self.assertIn('Balance', values)

        has_numeric_tb_rows = any(
            isinstance(cell.value, (int, float)) and abs(cell.value) > 0
            for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 300), min_col=1, max_col=min(ws.max_column, 12))
            for cell in row
        )
        self.assertTrue(has_numeric_tb_rows)

    def test_05_client_details_are_populated(self):
        wizard = self._create_wizard()
        _action, wb = self._export_workbook(wizard)

        ws = wb['Client Details']
        self.assertEqual(ws['B4'].value, self.env.company.name)
        self.assertEqual(self._coerce_openpyxl_date(ws['B5'].value), self.export_date_from)
        self.assertEqual(self._coerce_openpyxl_date(ws['B6'].value), self.export_date_to)

    def test_06_share_capital_rows_are_populated_when_company_data_exists(self):
        self.env.company.write({
            'shareholder_1': 'Owner A',
            'number_of_shares_1': 100,
            'share_value_1': 10.0,
        })

        wizard = self._create_wizard()
        _action, wb = self._export_workbook(wizard)

        ws = wb['Share Capital']
        self.assertEqual(ws['A7'].value, 'Owner A')
        self.assertEqual(ws['B7'].value, 100)
        self.assertEqual(ws['C7'].value, 10.0)

    def test_07_non_tb_template_sheets_have_no_static_numeric_samples(self):
        wizard = self._create_wizard()
        _action, wb = self._export_workbook(wizard)

        non_tb_template_sheets = [
            'VAT Control',
            'Prepayment',
            'Summary Sheet',
            'SOFP',
            'SOCI',
            'SOCE',
            'SOCF',
            'Accruals',
        ]

        for sheet_name in non_tb_template_sheets:
            ws = wb[sheet_name]
            static_values = self._static_numeric_like_values(ws)
            self.assertFalse(static_values, f'Static sample numeric values found on {sheet_name}: {static_values[:5]}')

    def test_08_trial_balance_forces_flat_unfolded_options(self):
        wizard = self._create_wizard(unfold_all=False)

        with patch.object(type(wizard), '_build_native_report_sheet_payload', autospec=True, return_value={}) as mocked_builder:
            wizard._prepare_trial_balance_payload()

        options = mocked_builder.call_args.kwargs['options']
        self.assertFalse(options.get('hierarchy'))
        self.assertTrue(options.get('unfold_all'))
        self.assertEqual(options.get('unfolded_lines'), [])

    def test_09_customer_invoice_report_options_are_forwarded(self):
        wizard = self._create_wizard(invoice_bill_scope='posted_only', include_refunds=False)

        with patch.object(type(wizard), '_get_invoice_bill_moves', autospec=True, return_value=self.env['account.move']) as mocked_get_moves:
            wizard._prepare_invoice_bill_payload(is_customer=True)

        options = mocked_get_moves.call_args.kwargs['options']
        self.assertEqual(options.get('invoice_bill_report_kind'), 'customer')
        self.assertEqual(options.get('invoice_bill_scope'), 'posted_only')
        self.assertFalse(options.get('include_refunds'))

    def test_10_statement_sheets_follow_pdf_layout_without_trial_balance_links(self):
        wizard = self._create_wizard()
        _action, wb = self._export_workbook(wizard)

        sofp = wb['SOFP']
        expected_sofp_labels = [
            'Assets:',
            'Non-current assets',
            'Total non-current assets',
            'Current assets',
            'Total current assets',
            'Total assets',
            'Equity and Liabilities:',
            'Equity',
            'Total Equity',
            'Non-current liabilities',
            'Current liabilities',
            'Total Liabilities',
            'Total Equity and Liabilities',
        ]
        sofp_rows = []
        for label in expected_sofp_labels:
            row_idx = self._find_row_by_label(sofp, label)
            self.assertIsNotNone(row_idx, f'Missing expected SOFP row: {label}')
            sofp_rows.append(row_idx)
        self.assertEqual(sofp_rows, sorted(sofp_rows), 'SOFP labels are out of expected order')

        unexpected_sofp_labels = [
            'Advances',
            'Prepaid expenses',
            'Bank overdraft',
            'Short-term loan',
            'Credit card payable',
            'Lease liability',
        ]
        for label in unexpected_sofp_labels:
            self.assertIsNone(self._find_row_by_label(sofp, label), f'Unexpected SOFP row present: {label}')

        def _assert_no_trial_balance_formulas(ws, sheet_name):
            references = []
            for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 400), min_col=1, max_col=min(ws.max_column, 12)):
                for cell in row:
                    value = cell.value
                    if isinstance(value, str) and value.startswith('=') and 'Trial Balance' in value:
                        references.append(cell.coordinate)
            self.assertFalse(references, f'Unexpected Trial Balance formula refs in {sheet_name}: {references[:5]}')

        for sheet_name in ('SOFP', 'SOCI', 'SOCE', 'SOCF'):
            _assert_no_trial_balance_formulas(wb[sheet_name], sheet_name)

    def test_11_last_saved_settings_are_loaded(self):
        default_key_wizard = self._create_wizard()
        settings_key = default_key_wizard._previous_settings_key()
        self.env['ir.config_parameter'].sudo().set_param(settings_key, '')

        custom_wizard = self._create_wizard(
            date_from=fields.Date.from_string('2025-02-01'),
            date_to=fields.Date.from_string('2025-02-28'),
            aged_as_of_date=fields.Date.from_string('2025-02-28'),
            export_sheet_key='general_ledger',
            audit_period_category='normal_2y',
            balance_sheet_date_mode='range',
            prior_year_mode='manual',
            prior_balance_sheet_date_mode='range',
            prior_date_start=fields.Date.from_string('2024-02-01'),
            prior_date_end=fields.Date.from_string('2024-02-29'),
            include_draft_entries=False,
            unfold_all=False,
            hide_zero_lines=True,
            aging_based_on='base_on_invoice_date',
            aging_interval=45,
            show_currency=False,
            show_account=False,
            include_dynamic_columns=False,
            invoice_bill_scope='posted_only',
            include_refunds=False,
            journal_ids=[Command.set(self.company_data['default_journal_sale'].ids)],
            partner_ids=[Command.set([self.customer_partner.id])],
            gl_options_json='{"unfold_all": false}',
            aged_receivable_options_json='{"show_currency": false}',
            aged_payable_options_json='{"show_currency": false}',
        )
        custom_wizard._store_previous_settings()

        field_names = [
            'use_previous_settings',
            'date_from',
            'date_to',
            'aged_as_of_date',
            'export_sheet_key',
            'audit_period_category',
            'balance_sheet_date_mode',
            'prior_year_mode',
            'prior_balance_sheet_date_mode',
            'prior_date_start',
            'prior_date_end',
            'include_draft_entries',
            'unfold_all',
            'hide_zero_lines',
            'aging_based_on',
            'aging_interval',
            'show_currency',
            'show_account',
            'include_dynamic_columns',
            'invoice_bill_scope',
            'include_refunds',
            'journal_ids',
            'partner_ids',
            'gl_options_json',
            'aged_receivable_options_json',
            'aged_payable_options_json',
        ]
        defaults = self.env['audit.excel.export.wizard'].default_get(field_names)

        self.assertEqual(defaults.get('date_from'), fields.Date.from_string('2025-02-01'))
        self.assertEqual(defaults.get('date_to'), fields.Date.from_string('2025-02-28'))
        self.assertEqual(defaults.get('aged_as_of_date'), fields.Date.from_string('2025-02-28'))
        self.assertEqual(defaults.get('export_sheet_key'), 'general_ledger')
        self.assertEqual(defaults.get('audit_period_category'), 'normal_2y')
        self.assertEqual(defaults.get('balance_sheet_date_mode'), 'range')
        self.assertEqual(defaults.get('prior_year_mode'), 'manual')
        self.assertEqual(defaults.get('prior_balance_sheet_date_mode'), 'range')
        self.assertEqual(defaults.get('prior_date_start'), fields.Date.from_string('2024-02-01'))
        self.assertEqual(defaults.get('prior_date_end'), fields.Date.from_string('2024-02-29'))
        self.assertFalse(defaults.get('include_draft_entries'))
        self.assertFalse(defaults.get('unfold_all'))
        self.assertTrue(defaults.get('hide_zero_lines'))
        self.assertEqual(defaults.get('aging_based_on'), 'base_on_invoice_date')
        self.assertEqual(defaults.get('aging_interval'), 45)
        self.assertFalse(defaults.get('show_currency'))
        self.assertFalse(defaults.get('show_account'))
        self.assertFalse(defaults.get('include_dynamic_columns'))
        self.assertEqual(defaults.get('invoice_bill_scope'), 'posted_only')
        self.assertFalse(defaults.get('include_refunds'))
        self.assertEqual(
            defaults.get('journal_ids'),
            [(6, 0, self.company_data['default_journal_sale'].ids)],
        )
        self.assertEqual(
            defaults.get('partner_ids'),
            [(6, 0, [self.customer_partner.id])],
        )
        self.assertEqual(defaults.get('gl_options_json'), '{"unfold_all": false}')
        self.assertEqual(defaults.get('aged_receivable_options_json'), '{"show_currency": false}')
        self.assertEqual(defaults.get('aged_payable_options_json'), '{"show_currency": false}')

    def test_12_can_export_only_general_ledger_sheet(self):
        wizard = self._create_wizard(export_sheet_key='general_ledger')
        _action, wb = self._export_workbook(wizard)

        self.assertEqual(wb.sheetnames, ['General Ledger'])
        values = self._sheet_values(wb['General Ledger'])
        self.assertIn('Code', values)
        self.assertIn('Account Name', values)
        self.assertIn('Debit', values)
        self.assertIn('Credit', values)
        self.assertIn('Balance', values)
        self.assertIn('general_ledger', wizard.file_name)

    def test_13_can_export_only_template_sheet(self):
        wizard = self._create_wizard(export_sheet_key='summary_sheet')
        _action, wb = self._export_workbook(wizard)

        self.assertEqual(wb.sheetnames, ['Summary Sheet'])
        self.assertTrue(str(wb['Summary Sheet']['A1'].value).startswith('='))

    def test_14_narration_parser_samples(self):
        parsed_1 = clean_bank_narration(
            'BNK1/2025/00042 INWARD T/T Return of deposit MARINA TSYBINA'
        )
        self.assertEqual(parsed_1.get('direction'), 'IN')
        self.assertEqual(parsed_1.get('rail'), 'TT')
        self.assertEqual(parsed_1.get('bank_ref'), 'BNK1/2025/00042')
        self.assertIn('IN TT', parsed_1.get('clean', ''))
        self.assertIn('BNK1/2025/00042', parsed_1.get('clean', ''))

        parsed_2 = clean_bank_narration(
            'BNK1/2025/00043 CHARGE COLLECTION-INCL. VAT '
            'Charge Collection-INCL. VAT OUTWARD REMITTANCE CHARGE'
        )
        self.assertEqual(parsed_2.get('rail'), 'FEE')
        self.assertIn('VAT', parsed_2.get('clean', ''))
        self.assertEqual(parsed_2.get('bank_ref'), 'BNK1/2025/00043')

        parsed_3 = clean_bank_narration(
            'BNK1/2025/00047 CREDIT CARD PAYMENT 536272XXXXXX7871'
        )
        self.assertEqual(parsed_3.get('rail'), 'CARD')
        self.assertEqual(parsed_3.get('card_last4'), '7871')
        self.assertIn('****7871', parsed_3.get('clean', ''))

        parsed_4 = clean_bank_narration(
            'BNK1/2025/00053 INWARD T/T /REF//BENEFRES/AE//PRS INVOICE NUMB ER 224 GOVINDER SINGH SIDHU'
        )
        self.assertEqual(parsed_4.get('direction'), 'IN')
        self.assertEqual(parsed_4.get('odoo_ref'), 'PRS-224')
        self.assertEqual(parsed_4.get('bank_ref'), 'BNK1/2025/00053')
        self.assertNotIn('1.0000', parsed_4.get('clean', ''))

        long_text = (
            'BNK1/2025/00999 INWARD T/T '
            + ('VERY LONG DESCRIPTION ' * 20)
            + 'PRS INVOICE NUMBER 998'
        )
        parsed_long = clean_bank_narration(long_text, max_len=100)
        self.assertLessEqual(len(parsed_long.get('clean', '')), 100)

    def test_15_clean_general_ledger_rows_and_keep_non_detail_rows(self):
        wizard = self._create_wizard()
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = 'General Ledger'

        sheet['A3'] = 'Code'
        sheet['B3'] = 'Account Name'
        sheet['G3'] = 'Debit'
        sheet['H3'] = 'Credit'
        sheet['I3'] = 'Balance'

        sheet['A4'] = '12020101'
        sheet['B4'] = 'Accounts Receivable'
        sheet['G4'] = 100.0
        sheet['H4'] = 0.0
        sheet['I4'] = 100.0

        sheet['B5'] = 'BNK1/2025/00047 CREDIT CARD PAYMENT 536272XXXXXX7871'
        sheet['G5'] = 0.0
        sheet['H5'] = 10.0
        sheet['I5'] = -10.0

        sheet['B6'] = 'Initial Balance'
        sheet['B7'] = 'Total 12020101 Account Receivable'
        sheet['B8'] = 'Load more...'
        sheet['B9'] = 'BNK1/2025/00053 INWARD T/T /REF//BENEFRES/AE//PRS INVOICE NUMB ER 224 GOVINDER SINGH SIDHU'

        before_amount_tuple = (sheet['G5'].value, sheet['H5'].value, sheet['I5'].value)
        wizard._clean_general_ledger_sheet_narrations(sheet)

        self.assertEqual(sheet['B4'].value, 'Accounts Receivable')
        self.assertEqual(sheet['B6'].value, 'Initial Balance')
        self.assertEqual(sheet['B7'].value, 'Total 12020101 Account Receivable')
        self.assertEqual(sheet['B8'].value, 'Load more...')
        self.assertTrue(str(sheet['B5'].value).startswith('CARD | Payment | ****7871'))
        self.assertIn('Raw narration:', sheet['B5'].comment.text)
        self.assertIn('CREDIT CARD PAYMENT 536272XXXXXX7871', sheet['B5'].comment.text)
        self.assertEqual(before_amount_tuple, (sheet['G5'].value, sheet['H5'].value, sheet['I5'].value))
        self.assertIn('PRS-224', str(sheet['B9'].value))

    def test_16_general_ledger_cleaner_feature_flag_in_stage_b(self):
        wizard = self._create_wizard(export_sheet_key='general_ledger')
        config = self.env['ir.config_parameter'].sudo()
        parameter_key = 'audit_excel_export.gl_narration_cleaner_enabled'

        source_workbook = Workbook()
        source_sheet = source_workbook.active
        source_sheet.title = 'General Ledger'
        source_sheet['A3'] = 'Code'
        source_sheet['B3'] = 'Account Name'
        source_sheet['B4'] = 'BNK1/2025/00047 CREDIT CARD PAYMENT 536272XXXXXX7871'

        stage_a_payload = {
            'data_sheets': [
                {
                    'sheet_name': 'General Ledger',
                    'native_source_sheet': source_sheet,
                },
            ],
            'trial_balance': {},
        }

        config.set_param(parameter_key, 'false')
        stage_b_disabled = wizard._stage_b_create_workbook_and_write_data_sheets(stage_a_payload)
        disabled_sheet = stage_b_disabled['workbook']['General Ledger']
        self.assertEqual(disabled_sheet['B4'].value, 'BNK1/2025/00047 CREDIT CARD PAYMENT 536272XXXXXX7871')
        self.assertFalse(disabled_sheet['B4'].comment)

        config.set_param(parameter_key, '')
        source_workbook_enabled = Workbook()
        source_sheet_enabled = source_workbook_enabled.active
        source_sheet_enabled.title = 'General Ledger'
        source_sheet_enabled['A3'] = 'Code'
        source_sheet_enabled['B3'] = 'Account Name'
        source_sheet_enabled['B4'] = 'BNK1/2025/00047 CREDIT CARD PAYMENT 536272XXXXXX7871'

        stage_b_enabled = wizard._stage_b_create_workbook_and_write_data_sheets(
            {
                'data_sheets': [
                    {
                        'sheet_name': 'General Ledger',
                        'native_source_sheet': source_sheet_enabled,
                    },
                ],
                'trial_balance': {},
            }
        )
        enabled_sheet = stage_b_enabled['workbook']['General Ledger']
        self.assertTrue(str(enabled_sheet['B4'].value).startswith('CARD | Payment | ****7871'))
        self.assertTrue(enabled_sheet['B4'].comment)

    def test_17_statement_export_fails_if_audit_report_data_is_unavailable(self):
        wizard = self._create_wizard(export_sheet_key='sofp')

        with patch.object(
            type(wizard),
            '_get_audit_report_statement_data',
            autospec=True,
            side_effect=RuntimeError('audit report fetch failed'),
        ):
            with self.assertRaises(ValidationError) as err:
                wizard.action_export_xlsx()

        self.assertIn('Unable to populate statement sheets from Audit Report data', str(err.exception))

    def test_18_one_year_period_category_skips_prior_period_windows(self):
        wizard = self._create_wizard(
            audit_period_category='normal_1y',
            prior_year_mode='manual',
            prior_date_start=False,
            prior_date_end=False,
        )

        periods = wizard._get_reporting_periods()
        self.assertFalse(periods.get('show_prior_year'))
        self.assertFalse(periods.get('prior_date_start'))
        self.assertFalse(periods.get('prior_date_end'))
        self.assertFalse(periods.get('prior_opening_date_end'))

    def test_19_one_year_sofp_hides_prior_year_column_data(self):
        wizard = self._create_wizard(
            export_sheet_key='sofp',
            audit_period_category='normal_1y',
            prior_year_mode='manual',
            prior_date_start=False,
            prior_date_end=False,
        )
        _action, wb = self._export_workbook(wizard)

        sofp = wb['SOFP']
        self.assertIsNone(sofp['C6'].value)
        for row_idx in range(7, sofp.max_row + 1):
            if sofp.cell(row=row_idx, column=1).value in (None, ''):
                continue
            self.assertIsNone(sofp[f'C{row_idx}'].value)

    def test_20_soci_uses_operating_expense_note_lines_dynamically(self):
        wizard = self._create_wizard(export_sheet_key='soci')

        dynamic_lines = [
            {
                'name': f'Expense line {idx:02d}',
                'current': float(idx),
                'prev': float(idx) / 2.0,
            }
            for idx in range(1, 13)
        ]
        stmt_data = {
            'rd': {
                'show_prior_year': True,
                'note_sections': [
                    {
                        'label': 'Operating expenses',
                        'lines': dynamic_lines,
                    },
                ],
            },
            'period_totals': {
                '4101': -1000.0,
                '4102': -100.0,
                '5101': 200.0,
                '4103': -50.0,
            },
            'prior_period_totals': {
                '4101': -900.0,
                '4102': -80.0,
                '5101': 180.0,
                '4103': -40.0,
            },
        }

        with patch.object(type(wizard), '_get_audit_report_statement_data', autospec=True, return_value=stmt_data):
            _action, wb = self._export_workbook(wizard)

        soci = wb['SOCI']
        total_operating_row = self._find_row_by_label(soci, 'Total operating expenses')
        self.assertIsNotNone(total_operating_row)
        self.assertGreater(total_operating_row, 27, 'Expected dynamic SOCI detail expansion')

        for idx in range(1, 13):
            label = f'Expense line {idx:02d}'
            row_idx = self._find_row_by_label(soci, label)
            self.assertIsNotNone(row_idx, f'Missing dynamic SOCI expense row: {label}')
            self.assertEqual(soci[f'B{row_idx}'].value, float(idx))
            self.assertEqual(soci[f'C{row_idx}'].value, float(idx) / 2.0)

    def test_21_soci_includes_investment_gain_loss_line(self):
        wizard = self._create_wizard(export_sheet_key='soci')

        stmt_data = {
            'rd': {
                'show_prior_year': True,
                'note_sections': [],
            },
            'period_totals': {
                '4101': -1000.0,
                '4102': -100.0,
                '5101': 200.0,
                '4103': -50.0,
                '5201': -25.0,
            },
            'prior_period_totals': {
                '4101': -900.0,
                '4102': -80.0,
                '5101': 180.0,
                '4103': -40.0,
                '5201': 10.0,
            },
        }

        with patch.object(type(wizard), '_get_audit_report_statement_data', autospec=True, return_value=stmt_data):
            _action, wb = self._export_workbook(wizard)

        soci = wb['SOCI']
        investment_row = self._find_row_by_label(soci, 'Gain / (loss) on investment')
        self.assertIsNotNone(investment_row)
        self.assertEqual(soci[f'B{investment_row}'].value, 25.0)
        self.assertEqual(soci[f'C{investment_row}'].value, -10.0)

    def test_22_soce_expands_rows_instead_of_truncating(self):
        wizard = self._create_wizard(export_sheet_key='soce')

        soce_rows = []
        for idx in range(1, 17):
            soce_rows.append({
                'label': f'SOCE Dynamic Row {idx:02d}',
                'share_capital': float(idx),
                'owner_current_account': float(idx + 10),
                'retained_earnings': float(idx + 20),
                'statutory_reserves': float(idx + 30),
                'total_equity': float(idx + 60),
                'is_balance': idx in (1, 8, 16),
            })

        stmt_data = {
            'rd': {
                'soce_rows': soce_rows,
                'show_prior_year': True,
            },
            'period_totals': {},
            'prior_period_totals': {},
        }
        with patch.object(type(wizard), '_get_audit_report_statement_data', autospec=True, return_value=stmt_data):
            _action, wb = self._export_workbook(wizard)

        soce = wb['SOCE']
        last_row_label = 'SOCE Dynamic Row 16'
        row_idx = self._find_row_by_label(soce, last_row_label)
        self.assertIsNotNone(row_idx, 'SOCE dynamic rows were truncated')
        self.assertGreater(row_idx, 19, 'Expected SOCE template expansion beyond static rows')
        self.assertEqual(soce[f'B{row_idx}'].value, 16.0)
        self.assertEqual(soce[f'F{row_idx}'].value, 76.0)

    def test_22_socf_dynamic_lines_and_section_hiding_follow_audit_logic(self):
        wizard = self._create_wizard(export_sheet_key='socf')

        stmt_data = {
            'rd': {
                'show_prior_year': True,
                'comparative_period_word': 'year',
                'cashflow_net_profit_amount': 100.0,
                'cashflow_prev_net_profit_amount': 80.0,
                'current_depreciation_total': 10.0,
                'prior_depreciation_total': 8.0,
                'end_service_benefits_adjustment': 5.0,
                'prior_end_service_benefits_adjustment': 4.0,
                'operating_cashflow_before_working_capital': 115.0,
                'prior_operating_cashflow_before_working_capital': 92.0,
                'change_in_current_assets': -20.0,
                'prior_change_in_current_assets': -10.0,
                'change_in_current_liabilities': 15.0,
                'prior_change_in_current_liabilities': 7.0,
                'corporate_tax_paid': -3.0,
                'prior_corporate_tax_paid': -2.0,
                'end_service_benefits_paid': -1.0,
                'prior_end_service_benefits_paid': 0.0,
                'net_cash_generated_from_operations': 106.0,
                'prior_net_cash_generated_from_operations': 87.0,
                'current_property': 0.0,
                'prior_property': 0.0,
                'net_cash_generated_from_investing_activities': 0.0,
                'prior_net_cash_generated_from_investing_activities': 0.0,
                'paid_up_capital': 0.0,
                'prior_paid_up_capital': 0.0,
                'dividend_paid': 0.0,
                'prior_dividend_paid': 0.0,
                'owner_current_account': 0.0,
                'prior_owner_current_account': 0.0,
                'show_owner_current_account_cashflow_row': False,
                'net_cash_generated_from_financing_activities': 0.0,
                'prior_net_cash_generated_from_financing_activities': 0.0,
                'net_cash_and_cash_equivalents': 106.0,
                'prior_net_cash_and_cash_equivalents': 87.0,
                'cash_beginning_year': 10.0,
                'prior_cash_beginning_year': 5.0,
                'cash_end_of_year': 116.0,
                'prior_cash_end_of_year': 92.0,
            },
            'period_totals': {},
            'prior_period_totals': {},
        }
        with patch.object(type(wizard), '_get_audit_report_statement_data', autospec=True, return_value=stmt_data):
            _action, wb = self._export_workbook(wizard)

        socf = wb['SOCF']
        corporate_tax_row = self._find_row_by_label(socf, 'Corporate tax paid')
        eosb_paid_row = self._find_row_by_label(socf, 'End of service benefits paid')
        self.assertIsNotNone(corporate_tax_row)
        self.assertIsNotNone(eosb_paid_row)
        self.assertEqual(socf[f'B{corporate_tax_row}'].value, -3.0)
        self.assertEqual(socf[f'C{corporate_tax_row}'].value, -2.0)
        self.assertEqual(socf[f'B{eosb_paid_row}'].value, -1.0)

        financing_header_row = self._find_row_by_label(socf, 'Cash flows from financing activities')
        self.assertIsNotNone(financing_header_row)
        self.assertTrue(bool(socf.row_dimensions[financing_header_row].hidden))

    def test_23_year_end_date_syncs_aged_as_of_date(self):
        year_end_date = fields.Date.from_string('2025-12-31')
        wizard = self._create_wizard(
            date_to=year_end_date,
            aged_as_of_date=fields.Date.from_string('2025-11-30'),
        )

        self.assertEqual(wizard.date_to, year_end_date)
        self.assertEqual(wizard.aged_as_of_date, year_end_date)

    def test_24_aged_payload_uses_year_end_date(self):
        year_end_date = fields.Date.from_string('2025-12-31')
        wizard = self._create_wizard(
            date_to=year_end_date,
            aged_as_of_date=fields.Date.from_string('2025-11-30'),
        )

        with patch.object(type(wizard), '_build_native_report_sheet_payload', autospec=True, return_value={}) as mocked_builder:
            wizard._prepare_aged_payload(
                xmlid='account_reports.aged_receivable_report',
                sheet_name='Aged Receivables',
                title='Aged Receivables',
                overrides={},
            )

        options = mocked_builder.call_args.kwargs['options']
        self.assertEqual(options.get('date', {}).get('date_from'), '2025-12-31')
        self.assertEqual(options.get('date', {}).get('date_to'), '2025-12-31')
