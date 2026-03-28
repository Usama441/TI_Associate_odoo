import csv
import json
import logging
import time
from dateutil.relativedelta import relativedelta
from odoo import models, fields, api
from odoo.exceptions import ValidationError

_logger = logging.getLogger(__name__)

CT_EXPENSE_ACCOUNT_CODES = ('51270101',)
CT_LIABILITY_ACCOUNT_CODES = ('22040101',)

class AuditReport(models.TransientModel):
    _name = 'audit.report'
    _description = 'Audit Report'

    SIGNATORY_ROLE_SELECTION = [
        ('primary', 'Primary'),
        ('secondary', 'Secondary'),
    ]

    date_start = fields.Date(string='Date Start', required=True)
    date_end = fields.Date(string='Date End', required=True)
    balance_sheet_date_mode = fields.Selection(
        [
            ('end_only', 'End date only (snapshot)'),
            ('range', 'Date range (start to end)'),
        ],
        string='Balance Sheet Dates',
        required=True,
        default='end_only',
    )
    company_id = fields.Many2one(
        'res.company',
        string='Company',
        required=True,
        default=lambda self: self.env.company,
    )
    company_name = fields.Char(related='company_id.name', readonly=False)
    company_street = fields.Char(related='company_id.street', readonly=False)
    company_street2 = fields.Char(related='company_id.street2', readonly=False)
    company_city = fields.Char(related='company_id.city', readonly=False)
    company_state_id = fields.Many2one(related='company_id.state_id', readonly=False)
    company_zip = fields.Char(related='company_id.zip', readonly=False)
    company_country_id = fields.Many2one(related='company_id.country_id', readonly=False)
    company_free_zone = fields.Selection(related='company_id.free_zone', readonly=False)
    company_license_number = fields.Char(related='company_id.company_license_number', readonly=False)
    trade_license_activities = fields.Text(related='company_id.trade_license_activities', readonly=False)
    incorporation_date = fields.Date(related='company_id.incorporation_date', readonly=False)
    corporate_tax_registration_number = fields.Char(
        related='company_id.corporate_tax_registration_number',
        readonly=False,
    )
    vat_registration_number = fields.Char(
        related='company_id.vat_registration_number',
        readonly=False,
    )
    corporate_tax_start_date = fields.Date(
        related='company_id.corporate_tax_start_date',
        readonly=False,
    )
    corporate_tax_end_date = fields.Date(
        related='company_id.corporate_tax_end_date',
        readonly=False,
    )
    implementing_regulations_freezone = fields.Text(
        related='company_id.implementing_regulations_freezone',
        readonly=False,
    )

    @api.onchange('company_free_zone')
    def _onchange_company_free_zone(self):
        if not self.company_free_zone or not self.company_id:
            return
        if hasattr(self.company_id, '_get_free_zone_implementing_regulations'):
            mapped = self.company_id._get_free_zone_implementing_regulations(self.company_free_zone)
            if mapped:
                self.implementing_regulations_freezone = mapped
    shareholder_1 = fields.Char(related='company_id.shareholder_1', readonly=False)
    shareholder_2 = fields.Char(related='company_id.shareholder_2', readonly=False)
    shareholder_3 = fields.Char(related='company_id.shareholder_3', readonly=False)
    shareholder_4 = fields.Char(related='company_id.shareholder_4', readonly=False)
    shareholder_5 = fields.Char(related='company_id.shareholder_5', readonly=False)
    shareholder_6 = fields.Char(related='company_id.shareholder_6', readonly=False)
    shareholder_7 = fields.Char(related='company_id.shareholder_7', readonly=False)
    shareholder_8 = fields.Char(related='company_id.shareholder_8', readonly=False)
    shareholder_9 = fields.Char(related='company_id.shareholder_9', readonly=False)
    shareholder_10 = fields.Char(related='company_id.shareholder_10', readonly=False)

    nationality_1 = fields.Char(related='company_id.nationality_1', readonly=False)
    nationality_2 = fields.Char(related='company_id.nationality_2', readonly=False)
    nationality_3 = fields.Char(related='company_id.nationality_3', readonly=False)
    nationality_4 = fields.Char(related='company_id.nationality_4', readonly=False)
    nationality_5 = fields.Char(related='company_id.nationality_5', readonly=False)
    nationality_6 = fields.Char(related='company_id.nationality_6', readonly=False)
    nationality_7 = fields.Char(related='company_id.nationality_7', readonly=False)
    nationality_8 = fields.Char(related='company_id.nationality_8', readonly=False)
    nationality_9 = fields.Char(related='company_id.nationality_9', readonly=False)
    nationality_10 = fields.Char(related='company_id.nationality_10', readonly=False)

    number_of_shares_1 = fields.Integer(related='company_id.number_of_shares_1', readonly=False)
    number_of_shares_2 = fields.Integer(related='company_id.number_of_shares_2', readonly=False)
    number_of_shares_3 = fields.Integer(related='company_id.number_of_shares_3', readonly=False)
    number_of_shares_4 = fields.Integer(related='company_id.number_of_shares_4', readonly=False)
    number_of_shares_5 = fields.Integer(related='company_id.number_of_shares_5', readonly=False)
    number_of_shares_6 = fields.Integer(related='company_id.number_of_shares_6', readonly=False)
    number_of_shares_7 = fields.Integer(related='company_id.number_of_shares_7', readonly=False)
    number_of_shares_8 = fields.Integer(related='company_id.number_of_shares_8', readonly=False)
    number_of_shares_9 = fields.Integer(related='company_id.number_of_shares_9', readonly=False)
    number_of_shares_10 = fields.Integer(related='company_id.number_of_shares_10', readonly=False)

    share_value_1 = fields.Float(related='company_id.share_value_1', readonly=False)
    share_value_2 = fields.Float(related='company_id.share_value_2', readonly=False)
    share_value_3 = fields.Float(related='company_id.share_value_3', readonly=False)
    share_value_4 = fields.Float(related='company_id.share_value_4', readonly=False)
    share_value_5 = fields.Float(related='company_id.share_value_5', readonly=False)
    share_value_6 = fields.Float(related='company_id.share_value_6', readonly=False)
    share_value_7 = fields.Float(related='company_id.share_value_7', readonly=False)
    share_value_8 = fields.Float(related='company_id.share_value_8', readonly=False)
    share_value_9 = fields.Float(related='company_id.share_value_9', readonly=False)
    share_value_10 = fields.Float(related='company_id.share_value_10', readonly=False)

    signature_include_1 = fields.Boolean(string='Signatory 1', default=False)
    signature_include_2 = fields.Boolean(string='Signatory 2', default=False)
    signature_include_3 = fields.Boolean(string='Signatory 3', default=False)
    signature_include_4 = fields.Boolean(string='Signatory 4', default=False)
    signature_include_5 = fields.Boolean(string='Signatory 5', default=False)
    signature_include_6 = fields.Boolean(string='Signatory 6', default=False)
    signature_include_7 = fields.Boolean(string='Signatory 7', default=False)
    signature_include_8 = fields.Boolean(string='Signatory 8', default=False)
    signature_include_9 = fields.Boolean(string='Signatory 9', default=False)
    signature_include_10 = fields.Boolean(string='Signatory 10', default=False)
    director_include_1 = fields.Boolean(string='Director 1', default=True)
    director_include_2 = fields.Boolean(string='Director 2', default=False)
    director_include_3 = fields.Boolean(string='Director 3', default=False)
    director_include_4 = fields.Boolean(string='Director 4', default=False)
    director_include_5 = fields.Boolean(string='Director 5', default=False)
    director_include_6 = fields.Boolean(string='Director 6', default=False)
    director_include_7 = fields.Boolean(string='Director 7', default=False)
    director_include_8 = fields.Boolean(string='Director 8', default=False)
    director_include_9 = fields.Boolean(string='Director 9', default=False)
    director_include_10 = fields.Boolean(string='Director 10', default=False)
    owner_include_1 = fields.Boolean(string='Owner 1', default=True)
    owner_include_2 = fields.Boolean(string='Owner 2', default=True)
    owner_include_3 = fields.Boolean(string='Owner 3', default=True)
    owner_include_4 = fields.Boolean(string='Owner 4', default=True)
    owner_include_5 = fields.Boolean(string='Owner 5', default=True)
    owner_include_6 = fields.Boolean(string='Owner 6', default=True)
    owner_include_7 = fields.Boolean(string='Owner 7', default=True)
    owner_include_8 = fields.Boolean(string='Owner 8', default=True)
    owner_include_9 = fields.Boolean(string='Owner 9', default=True)
    owner_include_10 = fields.Boolean(string='Owner 10', default=True)

    signature_role_1 = fields.Selection(
        SIGNATORY_ROLE_SELECTION,
        string='Signatory role 1',
        default='primary',
    )
    signature_role_2 = fields.Selection(
        SIGNATORY_ROLE_SELECTION,
        string='Signatory role 2',
        default='secondary',
    )
    signature_role_3 = fields.Selection(
        SIGNATORY_ROLE_SELECTION,
        string='Signatory role 3',
        default='secondary',
    )
    signature_role_4 = fields.Selection(
        SIGNATORY_ROLE_SELECTION,
        string='Signatory role 4',
        default='secondary',
    )
    signature_role_5 = fields.Selection(
        SIGNATORY_ROLE_SELECTION,
        string='Signatory role 5',
        default='secondary',
    )
    signature_role_6 = fields.Selection(
        SIGNATORY_ROLE_SELECTION,
        string='Signatory role 6',
        default='secondary',
    )
    signature_role_7 = fields.Selection(
        SIGNATORY_ROLE_SELECTION,
        string='Signatory role 7',
        default='secondary',
    )
    signature_role_8 = fields.Selection(
        SIGNATORY_ROLE_SELECTION,
        string='Signatory role 8',
        default='secondary',
    )
    signature_role_9 = fields.Selection(
        SIGNATORY_ROLE_SELECTION,
        string='Signatory role 9',
        default='secondary',
    )
    signature_role_10 = fields.Selection(
        SIGNATORY_ROLE_SELECTION,
        string='Signatory role 10',
        default='secondary',
    )

    total_share_1 = fields.Float(related='company_id.total_share_1', readonly=True)
    total_share_2 = fields.Float(related='company_id.total_share_2', readonly=True)
    total_share_3 = fields.Float(related='company_id.total_share_3', readonly=True)
    total_share_4 = fields.Float(related='company_id.total_share_4', readonly=True)
    total_share_5 = fields.Float(related='company_id.total_share_5', readonly=True)
    total_share_6 = fields.Float(related='company_id.total_share_6', readonly=True)
    total_share_7 = fields.Float(related='company_id.total_share_7', readonly=True)
    total_share_8 = fields.Float(related='company_id.total_share_8', readonly=True)
    total_share_9 = fields.Float(related='company_id.total_share_9', readonly=True)
    total_share_10 = fields.Float(related='company_id.total_share_10', readonly=True)
    auditor_type = fields.Selection(
        [
            ('default', 'Default'),
            ('ifza', 'IFZA'),
            ('dmcc', 'DMCC'),
        ],
        string='Freezone City',
        required=True,
        default='default',
    )
    report_type = fields.Selection(
        [
            ('period', 'Financial Statements for the Period Ended'),
            ('year', 'Financial Statements for the Year Ended'),
            ('management', 'Management Accounts for the Period Ended'),
        ],
        string='Report Type',
        required=True,
        default='period',
    )
    show_related_parties_note = fields.Boolean(
        string='Show related parties note',
        default=False,
    )
    show_shareholder_note = fields.Boolean(
        string='Show shareholder line and note',
        default=True,
    )
    corporate_tax_liability_paid = fields.Boolean(
        string='Corporate Tax Liability Paid',
        default=False,
    )
    show_ct_first_tax_year_line = fields.Boolean(
        string='Show CT first tax year line',
        default=True,
    )
    gap_report_of_directors = fields.Boolean(
        string='Gap in report of directors',
        default=True,
    )
    gap_independent_auditor_report = fields.Boolean(
        string='Gap in independent auditor report',
        default=True,
    )
    gap_notes_to_financial_statements = fields.Boolean(
        string='Gap in notes to financial statements',
        default=True,
    )
    signature_break_lines_report_of_directors = fields.Integer(
        string='Break lines (Report of Directors)',
        default=5,
    )
    signature_break_lines_balance_sheet = fields.Integer(
        string='Break lines (Statement of financial position)',
        default=5,
    )
    signature_break_lines_profit_loss = fields.Integer(
        string='Break lines (Statement of profit and loss)',
        default=5,
    )
    signature_break_lines_changes_in_equity = fields.Integer(
        string='Break lines (Statement of changes in equity)',
        default=5,
    )
    signature_break_lines_cash_flows = fields.Integer(
        string='Break lines (Statement of cash flows)',
        default=5,
    )
    signature_break_lines_notes = fields.Integer(
        string='Break lines (Notes to financial statements)',
        default=5,
    )
    use_previous_settings = fields.Boolean(
        string='Use Previous Settings',
        default=True,
    )
    audit_period_category = fields.Selection(
        [
            ('cessation_2y', '2 Years Cessation'),
            ('cessation_1y', '1 Year Cessation'),
            ('normal_1y', '1 Year Normal'),
            ('normal_2y', '2 Years Normal'),
            ('dormant_1y', '1 Year Dormant'),
            ('dormant_2y', '2 Years Dormant'),
        ],
        string='Audit Period Category',
        default='normal_2y',
    )
    share_capital_paid_status = fields.Selection(
        [
            ('paid', 'Paid'),
            ('unpaid', 'Unpaid'),
        ],
        string='Share capital status',
        default='paid',
    )
    show_share_capital_conversion_note = fields.Boolean(
        string='Show share capital conversion note',
        default=False,
    )
    share_conversion_currency = fields.Char(
        string='Original currency',
        default='GBP',
    )
    share_conversion_original_value = fields.Float(
        string='Original value per share',
        default=100.0,
    )
    share_conversion_exchange_rate = fields.Float(
        string='Exchange rate',
        default=4.66,
    )
    show_share_capital_transfer_note = fields.Boolean(
        string='Show share transfer note',
        default=False,
    )
    share_transfer_date = fields.Date(
        string='Transfer date',
    )
    share_transfer_from = fields.Char(
        string='Transferred from',
    )
    share_transfer_shares = fields.Integer(
        string='No. of shares transferred',
        default=0,
    )
    share_transfer_percentage = fields.Float(
        string='Transferred shares (%)',
        default=0.0,
    )
    share_transfer_to = fields.Char(
        string='Transferred to',
    )
    related_party_name = fields.Char(string='Related party')
    related_party_relationship = fields.Char(string='Nature of relationship')
    related_party_transaction = fields.Char(string='Nature of transactions')
    related_party_amount = fields.Float(string='Related party amount (current)')
    related_party_amount_prior = fields.Float(string='Related party amount (prior)')
    prior_year_mode = fields.Selection(
        [
            ('auto', 'Auto (1 year back)'),
            ('manual', 'Manual selection'),
        ],
        string='Prior Year Dates',
        required=True,
        default='auto',
    )
    prior_balance_sheet_date_mode = fields.Selection(
        [
            ('end_only', 'End date only (snapshot)'),
            ('range', 'Date range (start to end)'),
        ],
        string='Prior Balance Sheet Dates',
        required=True,
        default='end_only',
    )
    prior_date_start = fields.Date(string='Prior Date Start')
    prior_date_end = fields.Date(string='Prior Date End')
    soce_prior_opening_label_date = fields.Date(
        string='SOCE Prior Opening Date (label only)'
    )
    tb_override_line_ids = fields.One2many(
        'audit.report.tb.override.line',
        'wizard_id',
        string='Trial Balance Overrides',
        copy=False,
    )
    tb_include_zero_accounts = fields.Boolean(
        string='Include zero-balance accounts',
        default=False,
    )
    tb_overrides_json = fields.Text(
        string='TB Overrides Snapshot',
        copy=False,
    )
    tb_diff_current = fields.Float(
        string='Current mismatch',
        digits='Account',
        readonly=True,
        copy=False,
    )
    tb_diff_prior = fields.Float(
        string='Prior mismatch',
        digits='Account',
        readonly=True,
        copy=False,
    )
    tb_warning_current = fields.Char(
        string='Current warning',
        readonly=True,
        copy=False,
    )
    tb_warning_prior = fields.Char(
        string='Prior warning',
        readonly=True,
        copy=False,
    )
    emphasis_change_period = fields.Boolean(string='Change in Financial Year/Period')
    emphasis_correction_error = fields.Boolean(string='Correction of Error')
    emphasis_liquidation = fields.Boolean(string='Liquidation')
    emphasis_change_from_date = fields.Date(string='Changed From')
    emphasis_change_to_date = fields.Date(string='Changed To')
    additional_note_legal_status_change = fields.Boolean(string='Legal Status Change')
    additional_note_no_active_bank_account = fields.Boolean(string='No Active Bank Account')
    additional_note_business_bank_account_opened = fields.Boolean(
        string='Business Bank Account Opened Post Year-End'
    )
    emphasis_legal_change_date = fields.Date(string='Legal Change Date')
    emphasis_legal_status_from = fields.Char(string='Legal Status From')
    emphasis_legal_status_to = fields.Char(string='Legal Status To')

    def _previous_settings_key(self):
        return (
            f'audit_report_wizard.prev_settings.user_{self.env.user.id}.'
            f'company_{self.env.company.id}'
        )

    def _get_previous_settings(self):
        param = self.env['ir.config_parameter'].sudo()
        raw = param.get_param(self._previous_settings_key(), default='')
        if not raw:
            return {}
        try:
            return json.loads(raw)
        except (TypeError, ValueError):
            return {}

    def _store_previous_settings(self):
        data = {
            'date_start': self.date_start.isoformat() if self.date_start else False,
            'date_end': self.date_end.isoformat() if self.date_end else False,
            'balance_sheet_date_mode': self.balance_sheet_date_mode,
            'prior_year_mode': self.prior_year_mode,
            'prior_balance_sheet_date_mode': self.prior_balance_sheet_date_mode,
            'prior_date_start': self.prior_date_start.isoformat() if self.prior_date_start else False,
            'prior_date_end': self.prior_date_end.isoformat() if self.prior_date_end else False,
            'soce_prior_opening_label_date': (
                self.soce_prior_opening_label_date.isoformat()
                if self.soce_prior_opening_label_date else False
            ),
            'report_type': self.report_type,
            'auditor_type': self.auditor_type,
            'audit_period_category': self.audit_period_category,
            'show_related_parties_note': self.show_related_parties_note,
            'show_shareholder_note': self.show_shareholder_note,
            'corporate_tax_liability_paid': self.corporate_tax_liability_paid,
            'show_ct_first_tax_year_line': self.show_ct_first_tax_year_line,
            'gap_report_of_directors': self.gap_report_of_directors,
            'gap_independent_auditor_report': self.gap_independent_auditor_report,
            'gap_notes_to_financial_statements': self.gap_notes_to_financial_statements,
            'signature_break_lines_report_of_directors': self.signature_break_lines_report_of_directors,
            'signature_break_lines_balance_sheet': self.signature_break_lines_balance_sheet,
            'signature_break_lines_profit_loss': self.signature_break_lines_profit_loss,
            'signature_break_lines_changes_in_equity': self.signature_break_lines_changes_in_equity,
            'signature_break_lines_cash_flows': self.signature_break_lines_cash_flows,
            'signature_break_lines_notes': self.signature_break_lines_notes,
            'share_capital_paid_status': self.share_capital_paid_status,
            'show_share_capital_conversion_note': self.show_share_capital_conversion_note,
            'share_conversion_currency': self.share_conversion_currency,
            'share_conversion_original_value': self.share_conversion_original_value,
            'share_conversion_exchange_rate': self.share_conversion_exchange_rate,
            'show_share_capital_transfer_note': self.show_share_capital_transfer_note,
            'share_transfer_date': self.share_transfer_date.isoformat() if self.share_transfer_date else False,
            'share_transfer_from': self.share_transfer_from,
            'share_transfer_shares': self.share_transfer_shares,
            'share_transfer_percentage': self.share_transfer_percentage,
            'share_transfer_to': self.share_transfer_to,
            'related_party_name': self.related_party_name,
            'related_party_relationship': self.related_party_relationship,
            'related_party_transaction': self.related_party_transaction,
            'related_party_amount': self.related_party_amount,
            'related_party_amount_prior': self.related_party_amount_prior,
            'emphasis_change_period': self.emphasis_change_period,
            'emphasis_correction_error': self.emphasis_correction_error,
            'emphasis_liquidation': self.emphasis_liquidation,
            'emphasis_change_from_date': (
                self.emphasis_change_from_date.isoformat() if self.emphasis_change_from_date else False
            ),
            'emphasis_change_to_date': (
                self.emphasis_change_to_date.isoformat() if self.emphasis_change_to_date else False
            ),
            'additional_note_legal_status_change': self.additional_note_legal_status_change,
            'additional_note_no_active_bank_account': self.additional_note_no_active_bank_account,
            'additional_note_business_bank_account_opened': self.additional_note_business_bank_account_opened,
            'emphasis_legal_change_date': (
                self.emphasis_legal_change_date.isoformat() if self.emphasis_legal_change_date else False
            ),
            'emphasis_legal_status_from': self.emphasis_legal_status_from,
            'emphasis_legal_status_to': self.emphasis_legal_status_to,
        }
        for i in range(1, 11):
            include_key = f'signature_include_{i}'
            role_key = f'signature_role_{i}'
            director_key = f'director_include_{i}'
            owner_key = f'owner_include_{i}'
            data[include_key] = bool(getattr(self, include_key))
            data[role_key] = getattr(self, role_key) or ('primary' if i == 1 else 'secondary')
            data[director_key] = bool(getattr(self, director_key))
            data[owner_key] = bool(getattr(self, owner_key))
        self.env['ir.config_parameter'].sudo().set_param(
            self._previous_settings_key(),
            json.dumps(data),
        )

    @api.model
    def default_get(self, fields_list):
        res = super().default_get(fields_list)
        if not res.get('use_previous_settings', True):
            return res
        data = self._get_previous_settings()
        if not data:
            return res
        if 'date_start' in fields_list and data.get('date_start'):
            res['date_start'] = fields.Date.to_date(data['date_start'])
        if 'date_end' in fields_list and data.get('date_end'):
            res['date_end'] = fields.Date.to_date(data['date_end'])
        if 'balance_sheet_date_mode' in fields_list and data.get('balance_sheet_date_mode'):
            res['balance_sheet_date_mode'] = data['balance_sheet_date_mode']
        if 'prior_year_mode' in fields_list and data.get('prior_year_mode'):
            res['prior_year_mode'] = data['prior_year_mode']
        if 'prior_balance_sheet_date_mode' in fields_list:
            value = data.get('prior_balance_sheet_date_mode') or data.get('balance_sheet_date_mode')
            if value:
                res['prior_balance_sheet_date_mode'] = value
        if 'prior_date_start' in fields_list and data.get('prior_date_start'):
            res['prior_date_start'] = fields.Date.to_date(data['prior_date_start'])
        if 'prior_date_end' in fields_list and data.get('prior_date_end'):
            res['prior_date_end'] = fields.Date.to_date(data['prior_date_end'])
        if 'soce_prior_opening_label_date' in fields_list and data.get('soce_prior_opening_label_date'):
            res['soce_prior_opening_label_date'] = fields.Date.to_date(
                data['soce_prior_opening_label_date']
            )
        if 'report_type' in fields_list and data.get('report_type'):
            res['report_type'] = data['report_type']
        if 'auditor_type' in fields_list and data.get('auditor_type'):
            res['auditor_type'] = data['auditor_type']
        if 'audit_period_category' in fields_list and data.get('audit_period_category'):
            res['audit_period_category'] = data['audit_period_category']
        if 'show_related_parties_note' in fields_list and data.get('show_related_parties_note') is not None:
            res['show_related_parties_note'] = data['show_related_parties_note']
        if 'show_shareholder_note' in fields_list and data.get('show_shareholder_note') is not None:
            res['show_shareholder_note'] = data['show_shareholder_note']
        if 'corporate_tax_liability_paid' in fields_list and data.get('corporate_tax_liability_paid') is not None:
            res['corporate_tax_liability_paid'] = data['corporate_tax_liability_paid']
        if 'show_ct_first_tax_year_line' in fields_list and data.get('show_ct_first_tax_year_line') is not None:
            res['show_ct_first_tax_year_line'] = data['show_ct_first_tax_year_line']
        if 'gap_report_of_directors' in fields_list and data.get('gap_report_of_directors') is not None:
            res['gap_report_of_directors'] = data['gap_report_of_directors']
        if (
            'gap_independent_auditor_report' in fields_list
            and data.get('gap_independent_auditor_report') is not None
        ):
            res['gap_independent_auditor_report'] = data['gap_independent_auditor_report']
        if (
            'gap_notes_to_financial_statements' in fields_list
            and data.get('gap_notes_to_financial_statements') is not None
        ):
            res['gap_notes_to_financial_statements'] = data['gap_notes_to_financial_statements']
        old_directors_breaks = data.get('signature_break_lines_directors')
        old_other_breaks = data.get('signature_break_lines_other')
        if 'signature_break_lines_report_of_directors' in fields_list:
            value = data.get('signature_break_lines_report_of_directors')
            if value is None:
                value = old_directors_breaks if old_directors_breaks is not None else old_other_breaks
            if value is not None:
                res['signature_break_lines_report_of_directors'] = value
        if 'signature_break_lines_balance_sheet' in fields_list:
            value = data.get('signature_break_lines_balance_sheet')
            if value is None:
                value = old_other_breaks
            if value is not None:
                res['signature_break_lines_balance_sheet'] = value
        if 'signature_break_lines_profit_loss' in fields_list:
            value = data.get('signature_break_lines_profit_loss')
            if value is None:
                value = old_other_breaks
            if value is not None:
                res['signature_break_lines_profit_loss'] = value
        if 'signature_break_lines_changes_in_equity' in fields_list:
            value = data.get('signature_break_lines_changes_in_equity')
            if value is None:
                value = old_other_breaks
            if value is not None:
                res['signature_break_lines_changes_in_equity'] = value
        if 'signature_break_lines_cash_flows' in fields_list:
            value = data.get('signature_break_lines_cash_flows')
            if value is None:
                value = old_other_breaks
            if value is not None:
                res['signature_break_lines_cash_flows'] = value
        if 'signature_break_lines_notes' in fields_list:
            value = data.get('signature_break_lines_notes')
            if value is None:
                value = old_other_breaks
            if value is not None:
                res['signature_break_lines_notes'] = value
        if 'share_capital_paid_status' in fields_list and data.get('share_capital_paid_status'):
            res['share_capital_paid_status'] = data['share_capital_paid_status']
        if (
            'show_share_capital_conversion_note' in fields_list
            and data.get('show_share_capital_conversion_note') is not None
        ):
            res['show_share_capital_conversion_note'] = data['show_share_capital_conversion_note']
        if 'share_conversion_currency' in fields_list and data.get('share_conversion_currency') is not None:
            res['share_conversion_currency'] = data['share_conversion_currency']
        if (
            'share_conversion_original_value' in fields_list
            and data.get('share_conversion_original_value') is not None
        ):
            res['share_conversion_original_value'] = data['share_conversion_original_value']
        if (
            'share_conversion_exchange_rate' in fields_list
            and data.get('share_conversion_exchange_rate') is not None
        ):
            res['share_conversion_exchange_rate'] = data['share_conversion_exchange_rate']
        if (
            'show_share_capital_transfer_note' in fields_list
            and data.get('show_share_capital_transfer_note') is not None
        ):
            res['show_share_capital_transfer_note'] = data['show_share_capital_transfer_note']
        if 'share_transfer_date' in fields_list and data.get('share_transfer_date'):
            res['share_transfer_date'] = fields.Date.to_date(data['share_transfer_date'])
        if 'share_transfer_from' in fields_list and data.get('share_transfer_from') is not None:
            res['share_transfer_from'] = data['share_transfer_from']
        if 'share_transfer_shares' in fields_list and data.get('share_transfer_shares') is not None:
            res['share_transfer_shares'] = data['share_transfer_shares']
        if 'share_transfer_percentage' in fields_list and data.get('share_transfer_percentage') is not None:
            res['share_transfer_percentage'] = data['share_transfer_percentage']
        if 'share_transfer_to' in fields_list and data.get('share_transfer_to') is not None:
            res['share_transfer_to'] = data['share_transfer_to']
        if 'related_party_name' in fields_list and data.get('related_party_name'):
            res['related_party_name'] = data['related_party_name']
        if 'related_party_relationship' in fields_list and data.get('related_party_relationship'):
            res['related_party_relationship'] = data['related_party_relationship']
        if 'related_party_transaction' in fields_list and data.get('related_party_transaction'):
            res['related_party_transaction'] = data['related_party_transaction']
        if 'related_party_amount' in fields_list and data.get('related_party_amount') is not None:
            res['related_party_amount'] = data['related_party_amount']
        if 'related_party_amount_prior' in fields_list and data.get('related_party_amount_prior') is not None:
            res['related_party_amount_prior'] = data['related_party_amount_prior']
        if 'emphasis_change_period' in fields_list and data.get('emphasis_change_period') is not None:
            res['emphasis_change_period'] = data['emphasis_change_period']
        if 'emphasis_correction_error' in fields_list and data.get('emphasis_correction_error') is not None:
            res['emphasis_correction_error'] = data['emphasis_correction_error']
        if 'emphasis_liquidation' in fields_list and data.get('emphasis_liquidation') is not None:
            res['emphasis_liquidation'] = data['emphasis_liquidation']
        if 'emphasis_change_from_date' in fields_list and data.get('emphasis_change_from_date'):
            res['emphasis_change_from_date'] = fields.Date.to_date(data['emphasis_change_from_date'])
        if 'emphasis_change_to_date' in fields_list and data.get('emphasis_change_to_date'):
            res['emphasis_change_to_date'] = fields.Date.to_date(data['emphasis_change_to_date'])
        if (
            'additional_note_legal_status_change' in fields_list
            and data.get('additional_note_legal_status_change') is not None
        ):
            res['additional_note_legal_status_change'] = data['additional_note_legal_status_change']
        if (
            'additional_note_no_active_bank_account' in fields_list
            and data.get('additional_note_no_active_bank_account') is not None
        ):
            res['additional_note_no_active_bank_account'] = data['additional_note_no_active_bank_account']
        if (
            'additional_note_business_bank_account_opened' in fields_list
            and data.get('additional_note_business_bank_account_opened') is not None
        ):
            res['additional_note_business_bank_account_opened'] = data[
                'additional_note_business_bank_account_opened'
            ]
        if 'emphasis_legal_change_date' in fields_list and data.get('emphasis_legal_change_date'):
            res['emphasis_legal_change_date'] = fields.Date.to_date(data['emphasis_legal_change_date'])
        if 'emphasis_legal_status_from' in fields_list and data.get('emphasis_legal_status_from'):
            res['emphasis_legal_status_from'] = data['emphasis_legal_status_from']
        if 'emphasis_legal_status_to' in fields_list and data.get('emphasis_legal_status_to'):
            res['emphasis_legal_status_to'] = data['emphasis_legal_status_to']
        for i in range(1, 11):
            include_key = f'signature_include_{i}'
            role_key = f'signature_role_{i}'
            director_key = f'director_include_{i}'
            owner_key = f'owner_include_{i}'
            if include_key in fields_list and data.get(include_key) is not None:
                res[include_key] = data[include_key]
            if role_key in fields_list and data.get(role_key):
                res[role_key] = data[role_key]
            if director_key in fields_list and data.get(director_key) is not None:
                res[director_key] = data[director_key]
            if owner_key in fields_list and data.get(owner_key) is not None:
                res[owner_key] = data[owner_key]
        return res

    @api.onchange('use_previous_settings')
    def _onchange_use_previous_settings(self):
        if self.use_previous_settings:
            data = self._get_previous_settings()
            if data.get('date_start'):
                self.date_start = fields.Date.to_date(data['date_start'])
            if data.get('date_end'):
                self.date_end = fields.Date.to_date(data['date_end'])
            if data.get('balance_sheet_date_mode'):
                self.balance_sheet_date_mode = data['balance_sheet_date_mode']
            if data.get('prior_year_mode'):
                self.prior_year_mode = data['prior_year_mode']
            prior_balance_sheet_mode = (
                data.get('prior_balance_sheet_date_mode')
                or data.get('balance_sheet_date_mode')
            )
            if prior_balance_sheet_mode:
                self.prior_balance_sheet_date_mode = prior_balance_sheet_mode
            if data.get('prior_date_start'):
                self.prior_date_start = fields.Date.to_date(data['prior_date_start'])
            if data.get('prior_date_end'):
                self.prior_date_end = fields.Date.to_date(data['prior_date_end'])
            if data.get('soce_prior_opening_label_date'):
                self.soce_prior_opening_label_date = fields.Date.to_date(
                    data['soce_prior_opening_label_date']
                )
            if data.get('report_type'):
                self.report_type = data['report_type']
            if data.get('auditor_type'):
                self.auditor_type = data['auditor_type']
            if data.get('audit_period_category'):
                self.audit_period_category = data['audit_period_category']
            if data.get('show_related_parties_note') is not None:
                self.show_related_parties_note = data['show_related_parties_note']
            if data.get('show_shareholder_note') is not None:
                self.show_shareholder_note = data['show_shareholder_note']
            if data.get('corporate_tax_liability_paid') is not None:
                self.corporate_tax_liability_paid = data['corporate_tax_liability_paid']
            if data.get('show_ct_first_tax_year_line') is not None:
                self.show_ct_first_tax_year_line = data['show_ct_first_tax_year_line']
            if data.get('gap_report_of_directors') is not None:
                self.gap_report_of_directors = data['gap_report_of_directors']
            if data.get('gap_independent_auditor_report') is not None:
                self.gap_independent_auditor_report = data['gap_independent_auditor_report']
            if data.get('gap_notes_to_financial_statements') is not None:
                self.gap_notes_to_financial_statements = data['gap_notes_to_financial_statements']
            old_directors_breaks = data.get('signature_break_lines_directors')
            old_other_breaks = data.get('signature_break_lines_other')
            if data.get('signature_break_lines_report_of_directors') is not None:
                self.signature_break_lines_report_of_directors = data['signature_break_lines_report_of_directors']
            elif old_directors_breaks is not None:
                self.signature_break_lines_report_of_directors = old_directors_breaks
            elif old_other_breaks is not None:
                self.signature_break_lines_report_of_directors = old_other_breaks
            if data.get('signature_break_lines_balance_sheet') is not None:
                self.signature_break_lines_balance_sheet = data['signature_break_lines_balance_sheet']
            elif old_other_breaks is not None:
                self.signature_break_lines_balance_sheet = old_other_breaks
            if data.get('signature_break_lines_profit_loss') is not None:
                self.signature_break_lines_profit_loss = data['signature_break_lines_profit_loss']
            elif old_other_breaks is not None:
                self.signature_break_lines_profit_loss = old_other_breaks
            if data.get('signature_break_lines_changes_in_equity') is not None:
                self.signature_break_lines_changes_in_equity = data['signature_break_lines_changes_in_equity']
            elif old_other_breaks is not None:
                self.signature_break_lines_changes_in_equity = old_other_breaks
            if data.get('signature_break_lines_cash_flows') is not None:
                self.signature_break_lines_cash_flows = data['signature_break_lines_cash_flows']
            elif old_other_breaks is not None:
                self.signature_break_lines_cash_flows = old_other_breaks
            if data.get('signature_break_lines_notes') is not None:
                self.signature_break_lines_notes = data['signature_break_lines_notes']
            elif old_other_breaks is not None:
                self.signature_break_lines_notes = old_other_breaks
            if data.get('share_capital_paid_status'):
                self.share_capital_paid_status = data['share_capital_paid_status']
            if data.get('show_share_capital_conversion_note') is not None:
                self.show_share_capital_conversion_note = data['show_share_capital_conversion_note']
            if data.get('share_conversion_currency') is not None:
                self.share_conversion_currency = data['share_conversion_currency']
            if data.get('share_conversion_original_value') is not None:
                self.share_conversion_original_value = data['share_conversion_original_value']
            if data.get('share_conversion_exchange_rate') is not None:
                self.share_conversion_exchange_rate = data['share_conversion_exchange_rate']
            if data.get('show_share_capital_transfer_note') is not None:
                self.show_share_capital_transfer_note = data['show_share_capital_transfer_note']
            if data.get('share_transfer_date'):
                self.share_transfer_date = fields.Date.to_date(data['share_transfer_date'])
            if data.get('share_transfer_from') is not None:
                self.share_transfer_from = data['share_transfer_from']
            if data.get('share_transfer_shares') is not None:
                self.share_transfer_shares = data['share_transfer_shares']
            if data.get('share_transfer_percentage') is not None:
                self.share_transfer_percentage = data['share_transfer_percentage']
            if data.get('share_transfer_to') is not None:
                self.share_transfer_to = data['share_transfer_to']
            if data.get('related_party_name'):
                self.related_party_name = data['related_party_name']
            if data.get('related_party_relationship'):
                self.related_party_relationship = data['related_party_relationship']
            if data.get('related_party_transaction'):
                self.related_party_transaction = data['related_party_transaction']
            if data.get('related_party_amount') is not None:
                self.related_party_amount = data['related_party_amount']
            if data.get('related_party_amount_prior') is not None:
                self.related_party_amount_prior = data['related_party_amount_prior']
            if data.get('emphasis_change_period') is not None:
                self.emphasis_change_period = data['emphasis_change_period']
            if data.get('emphasis_correction_error') is not None:
                self.emphasis_correction_error = data['emphasis_correction_error']
            if data.get('emphasis_liquidation') is not None:
                self.emphasis_liquidation = data['emphasis_liquidation']
            if data.get('emphasis_change_from_date'):
                self.emphasis_change_from_date = fields.Date.to_date(data['emphasis_change_from_date'])
            if data.get('emphasis_change_to_date'):
                self.emphasis_change_to_date = fields.Date.to_date(data['emphasis_change_to_date'])
            if data.get('additional_note_legal_status_change') is not None:
                self.additional_note_legal_status_change = data['additional_note_legal_status_change']
            if data.get('additional_note_no_active_bank_account') is not None:
                self.additional_note_no_active_bank_account = data['additional_note_no_active_bank_account']
            if data.get('additional_note_business_bank_account_opened') is not None:
                self.additional_note_business_bank_account_opened = data[
                    'additional_note_business_bank_account_opened'
                ]
            if data.get('emphasis_legal_change_date'):
                self.emphasis_legal_change_date = fields.Date.to_date(data['emphasis_legal_change_date'])
            if data.get('emphasis_legal_status_from'):
                self.emphasis_legal_status_from = data['emphasis_legal_status_from']
            if data.get('emphasis_legal_status_to'):
                self.emphasis_legal_status_to = data['emphasis_legal_status_to']
            for i in range(1, 11):
                include_key = f'signature_include_{i}'
                role_key = f'signature_role_{i}'
                director_key = f'director_include_{i}'
                owner_key = f'owner_include_{i}'
                if data.get(include_key) is not None:
                    setattr(self, include_key, data[include_key])
                if data.get(role_key):
                    setattr(self, role_key, data[role_key])
                if data.get(director_key) is not None:
                    setattr(self, director_key, data[director_key])
                if data.get(owner_key) is not None:
                    setattr(self, owner_key, data[owner_key])
        else:
            self.date_start = False
            self.date_end = False
            self.balance_sheet_date_mode = 'end_only'
            self.prior_year_mode = 'auto'
            self.prior_balance_sheet_date_mode = 'end_only'
            self.prior_date_start = False
            self.prior_date_end = False
            self.soce_prior_opening_label_date = False
            self.report_type = 'period'
            self.auditor_type = 'default'
            self.audit_period_category = 'normal_2y'
            self.show_related_parties_note = False
            self.show_shareholder_note = True
            self.corporate_tax_liability_paid = False
            self.show_ct_first_tax_year_line = True
            self.gap_report_of_directors = True
            self.gap_independent_auditor_report = True
            self.gap_notes_to_financial_statements = True
            self.signature_break_lines_report_of_directors = 5
            self.signature_break_lines_balance_sheet = 5
            self.signature_break_lines_profit_loss = 5
            self.signature_break_lines_changes_in_equity = 5
            self.signature_break_lines_cash_flows = 5
            self.signature_break_lines_notes = 5
            self.share_capital_paid_status = 'paid'
            self.show_share_capital_conversion_note = False
            self.share_conversion_currency = 'GBP'
            self.share_conversion_original_value = 100.0
            self.share_conversion_exchange_rate = 4.66
            self.show_share_capital_transfer_note = False
            self.share_transfer_date = False
            self.share_transfer_from = False
            self.share_transfer_shares = 0
            self.share_transfer_percentage = 0.0
            self.share_transfer_to = False
            self.related_party_name = False
            self.related_party_relationship = False
            self.related_party_transaction = False
            self.related_party_amount = 0.0
            self.related_party_amount_prior = 0.0
            self.emphasis_change_period = False
            self.emphasis_correction_error = False
            self.emphasis_liquidation = False
            self.emphasis_change_from_date = False
            self.emphasis_change_to_date = False
            self.additional_note_legal_status_change = False
            self.additional_note_no_active_bank_account = False
            self.additional_note_business_bank_account_opened = False
            self.emphasis_legal_change_date = False
            self.emphasis_legal_status_from = False
            self.emphasis_legal_status_to = False
            self.tb_include_zero_accounts = False
            self.tb_override_line_ids = [(5, 0, 0)]
            self.tb_overrides_json = False
            self.tb_diff_current = 0.0
            self.tb_diff_prior = 0.0
            self.tb_warning_current = False
            self.tb_warning_prior = False
            self.company_id = self.env.company
            for i in range(1, 11):
                setattr(self, f'signature_include_{i}', False)
                setattr(self, f'signature_role_{i}', 'primary' if i == 1 else 'secondary')
                setattr(self, f'director_include_{i}', i == 1)
                setattr(self, f'owner_include_{i}', True)

    @api.onchange('emphasis_change_period')
    def _onchange_emphasis_change_period(self):
        if not self.emphasis_change_period:
            self.emphasis_change_from_date = False
            self.emphasis_change_to_date = False

    @api.onchange('additional_note_legal_status_change')
    def _onchange_additional_note_legal_status_change(self):
        if not self.additional_note_legal_status_change:
            self.emphasis_legal_change_date = False
            self.emphasis_legal_status_from = False
            self.emphasis_legal_status_to = False

    def _validate_emphasis_options(self):
        self.ensure_one()
        if self.emphasis_change_period:
            if not self.emphasis_change_from_date or not self.emphasis_change_to_date:
                raise ValidationError(
                    "Please provide both 'Changed From' and 'Changed To' dates for Emphasis of Matter."
                )
            if self.emphasis_change_from_date >= self.emphasis_change_to_date:
                raise ValidationError(
                    "'Changed To' date must be after 'Changed From' date for Emphasis of Matter."
                )
        if self.additional_note_legal_status_change:
            if (
                not self.emphasis_legal_change_date
                or not self.emphasis_legal_status_from
                or not self.emphasis_legal_status_to
            ):
                raise ValidationError(
                    "Please provide Legal Change Date, Legal Status From, and Legal Status To for Additional Notes."
                )
        if self.show_share_capital_conversion_note:
            currency = (self.share_conversion_currency or '').strip()
            if not currency:
                raise ValidationError(
                    "Please provide Original Currency for the share capital conversion note."
                )
            if (self.share_conversion_original_value or 0.0) <= 0.0:
                raise ValidationError(
                    "Original Value per Share must be greater than zero for the share capital conversion note."
                )
            if (self.share_conversion_exchange_rate or 0.0) <= 0.0:
                raise ValidationError(
                    "Exchange Rate must be greater than zero for the share capital conversion note."
                )
        if self.show_share_capital_transfer_note:
            if not self.share_transfer_date:
                raise ValidationError(
                    "Please provide Transfer Date for the share transfer note."
                )
            if not (self.share_transfer_from or '').strip():
                raise ValidationError(
                    "Please provide Transferred From for the share transfer note."
                )
            if (self.share_transfer_shares or 0) <= 0:
                raise ValidationError(
                    "No. of Shares Transferred must be greater than zero for the share transfer note."
                )
            percentage = self.share_transfer_percentage or 0.0
            if percentage <= 0.0:
                raise ValidationError(
                    "Transferred Shares (%) must be greater than zero for the share transfer note."
                )
            if percentage > 100.0:
                raise ValidationError(
                    "Transferred Shares (%) cannot be more than 100 for the share transfer note."
                )
            if not (self.share_transfer_to or '').strip():
                raise ValidationError(
                    "Please provide Transferred To for the share transfer note."
                )

    @api.onchange(
        'company_street',
        'company_free_zone',
        'company_license_number',
        'trade_license_activities',
        'incorporation_date',
        'corporate_tax_registration_number',
        'vat_registration_number',
        'corporate_tax_start_date',
        'corporate_tax_end_date',
        'implementing_regulations_freezone',
        'shareholder_1',
        'shareholder_2',
        'shareholder_3',
        'shareholder_4',
        'shareholder_5',
        'shareholder_6',
        'shareholder_7',
        'shareholder_8',
        'shareholder_9',
        'shareholder_10',
        'nationality_1',
        'nationality_2',
        'nationality_3',
        'nationality_4',
        'nationality_5',
        'nationality_6',
        'nationality_7',
        'nationality_8',
        'nationality_9',
        'nationality_10',
        'number_of_shares_1',
        'number_of_shares_2',
        'number_of_shares_3',
        'number_of_shares_4',
        'number_of_shares_5',
        'number_of_shares_6',
        'number_of_shares_7',
        'number_of_shares_8',
        'number_of_shares_9',
        'number_of_shares_10',
        'share_value_1',
        'share_value_2',
        'share_value_3',
        'share_value_4',
        'share_value_5',
        'share_value_6',
        'share_value_7',
        'share_value_8',
        'share_value_9',
        'share_value_10',
    )
    def _onchange_sync_company_info(self):
        if not self.company_id:
            return
        vals = {
            'street': self.company_street or False,
            'free_zone': self.company_free_zone or False,
            'company_license_number': self.company_license_number or False,
            'trade_license_activities': self.trade_license_activities or False,
            'incorporation_date': self.incorporation_date or False,
            'corporate_tax_registration_number': self.corporate_tax_registration_number or False,
            'vat_registration_number': self.vat_registration_number or False,
            'corporate_tax_start_date': self.corporate_tax_start_date or False,
            'corporate_tax_end_date': self.corporate_tax_end_date or False,
            'implementing_regulations_freezone': self.implementing_regulations_freezone or False,
            'shareholder_1': self.shareholder_1 or False,
            'shareholder_2': self.shareholder_2 or False,
            'shareholder_3': self.shareholder_3 or False,
            'shareholder_4': self.shareholder_4 or False,
            'shareholder_5': self.shareholder_5 or False,
            'shareholder_6': self.shareholder_6 or False,
            'shareholder_7': self.shareholder_7 or False,
            'shareholder_8': self.shareholder_8 or False,
            'shareholder_9': self.shareholder_9 or False,
            'shareholder_10': self.shareholder_10 or False,
            'nationality_1': self.nationality_1 or False,
            'nationality_2': self.nationality_2 or False,
            'nationality_3': self.nationality_3 or False,
            'nationality_4': self.nationality_4 or False,
            'nationality_5': self.nationality_5 or False,
            'nationality_6': self.nationality_6 or False,
            'nationality_7': self.nationality_7 or False,
            'nationality_8': self.nationality_8 or False,
            'nationality_9': self.nationality_9 or False,
            'nationality_10': self.nationality_10 or False,
            'number_of_shares_1': self.number_of_shares_1 or 0,
            'number_of_shares_2': self.number_of_shares_2 or 0,
            'number_of_shares_3': self.number_of_shares_3 or 0,
            'number_of_shares_4': self.number_of_shares_4 or 0,
            'number_of_shares_5': self.number_of_shares_5 or 0,
            'number_of_shares_6': self.number_of_shares_6 or 0,
            'number_of_shares_7': self.number_of_shares_7 or 0,
            'number_of_shares_8': self.number_of_shares_8 or 0,
            'number_of_shares_9': self.number_of_shares_9 or 0,
            'number_of_shares_10': self.number_of_shares_10 or 0,
            'share_value_1': self.share_value_1 or 0.0,
            'share_value_2': self.share_value_2 or 0.0,
            'share_value_3': self.share_value_3 or 0.0,
            'share_value_4': self.share_value_4 or 0.0,
            'share_value_5': self.share_value_5 or 0.0,
            'share_value_6': self.share_value_6 or 0.0,
            'share_value_7': self.share_value_7 or 0.0,
            'share_value_8': self.share_value_8 or 0.0,
            'share_value_9': self.share_value_9 or 0.0,
            'share_value_10': self.share_value_10 or 0.0,
        }
        self.company_id.write(vals)

    def _get_reporting_periods(self):
        self.ensure_one()
        date_start = self.date_start
        date_end = self.date_end
        period_category = (self.audit_period_category or '').lower()
        show_prior_year = not period_category.endswith('_1y')
        prior_year_mode = self.prior_year_mode
        prior_balance_sheet_date_mode = self.prior_balance_sheet_date_mode or 'end_only'
        prior_date_start_raw = self.prior_date_start
        prior_date_end_raw = self.prior_date_end

        if prior_year_mode == 'manual' and prior_date_start_raw and prior_date_end_raw:
            prior_date_start = fields.Date.to_date(prior_date_start_raw)
            prior_date_end = fields.Date.to_date(prior_date_end_raw)
        else:
            prior_date_start = date_start - relativedelta(years=1) if date_start else False
            prior_date_end = date_end - relativedelta(years=1) if date_end else False
        if prior_balance_sheet_date_mode == 'end_only' and prior_date_end:
            prior_date_start = prior_date_end - relativedelta(years=1) + relativedelta(days=1)

        prior_prior_date_start = (
            prior_date_start - relativedelta(years=1)
            if prior_date_start else False
        )
        prior_prior_date_end = (
            prior_date_end - relativedelta(years=1)
            if prior_date_end else False
        )

        balance_sheet_date_start = date_start if self.balance_sheet_date_mode == 'range' else False
        prior_balance_sheet_date_start = (
            prior_date_start if prior_balance_sheet_date_mode == 'range' else False
        )
        prior_prior_balance_sheet_date_start = (
            prior_prior_date_start if prior_balance_sheet_date_mode == 'range' else False
        )

        return {
            'date_start': date_start,
            'date_end': date_end,
            'period_category': period_category,
            'show_prior_year': show_prior_year,
            'prior_year_mode': prior_year_mode,
            'prior_balance_sheet_date_mode': prior_balance_sheet_date_mode,
            'prior_date_start': prior_date_start,
            'prior_date_end': prior_date_end,
            'prior_prior_date_start': prior_prior_date_start,
            'prior_prior_date_end': prior_prior_date_end,
            'balance_sheet_date_start': balance_sheet_date_start,
            'prior_balance_sheet_date_start': prior_balance_sheet_date_start,
            'prior_prior_balance_sheet_date_start': prior_prior_balance_sheet_date_start,
        }

    def _fetch_grouped_account_rows(self, date_start, date_end):
        self.ensure_one()
        if not date_end:
            return []

        domain = [
            ('date', '<=', date_end),
            ('company_id', '=', self.company_id.id),
            ('parent_state', '=', 'posted'),
        ]
        if date_start:
            domain.insert(0, ('date', '>=', date_start))

        move_line_env = self.env['account.move.line'].with_context(
            allowed_company_ids=[self.company_id.id]
        )
        grouped_rows = move_line_env._read_group(
            domain=domain,
            groupby=['account_id'],
            aggregates=['debit:sum', 'credit:sum', 'balance:sum'],
        )

        account_ids = [account.id for account, _debit, _credit, _balance in grouped_rows if account]
        account_env = self.env['account.account'].with_company(self.company_id).with_context(
            allowed_company_ids=[self.company_id.id]
        )
        account_map = {acc.id: acc for acc in account_env.browse(account_ids)}

        final_rows = []
        for account_row, debit_sum, credit_sum, balance_sum in grouped_rows:
            if not account_row:
                continue
            account_id = account_row.id
            account = account_map.get(account_id)
            if not account:
                continue
            code_raw = (account.code or account.code_store or '').strip()
            code = self._normalize_account_code(code_raw)
            if not code:
                continue
            final_rows.append({
                'id': account_id,
                'code': code,
                'code_raw': code_raw,
                'name': account.name,
                'type': account.account_type,
                'debit': debit_sum or 0.0,
                'credit': credit_sum or 0.0,
                'balance': balance_sum or 0.0,
            })

        return final_rows

    def _tb_identity_keys(self, account_id=False, account_code=False):
        keys = []
        if account_id:
            keys.append(('id', int(account_id)))
        normalized_code = self._normalize_account_code(account_code)
        if normalized_code:
            keys.append(('code', normalized_code))
        return keys

    def _serialize_tb_overrides_payload(self):
        self.ensure_one()
        payload = []
        ordered_lines = self.tb_override_line_ids.sorted(
            key=lambda line: (
                line.period_key or '',
                line.account_code or '',
                line.id,
            )
        )
        for line in ordered_lines:
            if not line.is_overridden:
                continue
            payload.append({
                'period_key': line.period_key,
                'account_id': line.account_id.id if line.account_id else False,
                'account_code': self._normalize_account_code(line.account_code),
                'account_name': line.account_name or '',
                'system_debit': self._to_float(line.system_debit),
                'system_credit': self._to_float(line.system_credit),
                'system_balance': self._to_float(line.system_balance),
                'override_debit': self._to_float(line.override_debit),
                'override_credit': self._to_float(line.override_credit),
                'override_balance': (
                    self._to_float(line.override_debit)
                    - self._to_float(line.override_credit)
                ),
            })
        return payload

    def _sync_tb_overrides_json(self):
        self.ensure_one()
        payload = self._serialize_tb_overrides_payload()
        serialized = json.dumps(payload)
        if (self.tb_overrides_json or '') != serialized:
            self.tb_overrides_json = serialized
        return serialized

    def _build_tb_override_maps(self):
        self.ensure_one()
        override_maps = {
            'current': {},
            'prior': {},
        }
        for line in self.tb_override_line_ids:
            period_key = line.period_key
            if period_key not in override_maps:
                continue
            if not line.is_overridden:
                continue
            effective_debit = self._to_float(line.override_debit)
            effective_credit = self._to_float(line.override_credit)
            payload = {
                'debit': effective_debit,
                'credit': effective_credit,
                'balance': effective_debit - effective_credit,
            }
            for identity_key in self._tb_identity_keys(
                line.account_id.id if line.account_id else False,
                line.account_code,
            ):
                override_maps[period_key][identity_key] = payload
        return override_maps

    @api.model
    def _apply_tb_overrides_to_rows(self, rows, override_map):
        if not override_map:
            return [dict(row) for row in rows]
        adjusted_rows = []
        for row in rows:
            adjusted_row = dict(row)
            override_payload = None
            row_id = adjusted_row.get('id')
            row_code = adjusted_row.get('code')
            if row_id:
                override_payload = override_map.get(('id', int(row_id)))
            if not override_payload and row_code:
                override_payload = override_map.get(('code', row_code))
            if override_payload:
                adjusted_row['debit'] = self._to_float(override_payload.get('debit'))
                adjusted_row['credit'] = self._to_float(override_payload.get('credit'))
                adjusted_row['balance'] = (
                    self._to_float(adjusted_row['debit']) - self._to_float(adjusted_row['credit'])
                )
            adjusted_rows.append(adjusted_row)
        return adjusted_rows

    @api.model
    def _is_display_non_zero(self, value, precision=2):
        return round(value or 0.0, precision) != 0.0

    @api.model
    def _compose_tb_warning(self, label, diff_value):
        if not self._is_display_non_zero(diff_value, precision=2):
            return False
        return (
            f"{label} mismatch: Total assets differ from total equity and liabilities by "
            f"{diff_value:,.2f}."
        )

    def _reopen_wizard_form(self):
        self.ensure_one()
        action = self.env['ir.actions.actions']._for_xml_id('Audit_Report.audit_report_wizard_action')
        action.update({
            'res_id': self.id,
            'view_mode': 'form',
            'target': 'new',
        })
        return action

    def _load_tb_override_lines(self, preserve_overrides=False):
        self.ensure_one()
        periods = self._get_reporting_periods()
        date_end = periods['date_end']
        show_prior_year = periods['show_prior_year']
        prior_date_end = periods['prior_date_end']
        balance_sheet_date_start = periods['balance_sheet_date_start']
        prior_balance_sheet_date_start = periods['prior_balance_sheet_date_start']
        include_zero_accounts = bool(self.tb_include_zero_accounts)

        target_periods = ['current']
        if show_prior_year:
            target_periods.append('prior')

        if not preserve_overrides:
            self.tb_override_line_ids.filtered(
                lambda line: line.period_key in target_periods
            ).unlink()

        existing_lines = self.tb_override_line_ids.filtered(
            lambda line: line.period_key in target_periods
        )
        lines_by_period = {period_key: {} for period_key in target_periods}
        for line in existing_lines:
            for identity_key in self._tb_identity_keys(
                line.account_id.id if line.account_id else False,
                line.account_code,
            ):
                lines_by_period.setdefault(line.period_key, {})
                lines_by_period[line.period_key][identity_key] = line

        period_specs = [
            ('current', balance_sheet_date_start, date_end),
        ]
        if show_prior_year:
            period_specs.append(('prior', prior_balance_sheet_date_start, prior_date_end))

        seen_line_ids = set()
        for period_key, period_start, period_end in period_specs:
            rows = self._fetch_grouped_account_rows(period_start, period_end)
            if not include_zero_accounts:
                rows = [
                    row for row in rows
                    if any(self._is_display_non_zero(row.get(field_name), precision=2)
                           for field_name in ('debit', 'credit', 'balance'))
                ]
            for row in rows:
                account_id = row.get('id')
                account_code = row.get('code')
                account_name = row.get('name') or ''
                system_debit = self._to_float(row.get('debit'))
                system_credit = self._to_float(row.get('credit'))
                system_balance = system_debit - system_credit
                identity_keys = self._tb_identity_keys(account_id, account_code)
                line = None
                for identity_key in identity_keys:
                    line = lines_by_period.get(period_key, {}).get(identity_key)
                    if line:
                        break

                system_vals = {
                    'account_id': account_id or False,
                    'account_code': account_code,
                    'account_name': account_name,
                    'system_debit': system_debit,
                    'system_credit': system_credit,
                    'system_balance': system_balance,
                }

                if line:
                    line.write(system_vals)
                    if not preserve_overrides:
                        line.write({
                            'override_debit': system_debit,
                            'override_credit': system_credit,
                            'override_balance': system_balance,
                        })
                    seen_line_ids.add(line.id)
                    continue

                create_vals = {
                    'wizard_id': self.id,
                    'period_key': period_key,
                    **system_vals,
                    'override_debit': system_debit,
                    'override_credit': system_credit,
                    'override_balance': system_balance,
                }
                created_line = self.env['audit.report.tb.override.line'].create(create_vals)
                seen_line_ids.add(created_line.id)

        if preserve_overrides:
            stale_lines = self.tb_override_line_ids.filtered(
                lambda line: (
                    line.period_key in target_periods
                    and line.id not in seen_line_ids
                    and not line.is_overridden
                )
            )
            stale_lines.unlink()

        if not show_prior_year:
            self.tb_override_line_ids.filtered(lambda line: line.period_key == 'prior').unlink()

        self._sync_tb_overrides_json()
        try:
            self._get_report_data()
        except Exception as err:
            _logger.debug(
                "Trial balance warning refresh skipped for wizard_id=%s due to: %s",
                self.id,
                err,
            )

    def action_tb_load_lines(self):
        self.ensure_one()
        self._load_tb_override_lines(preserve_overrides=False)
        return self._reopen_wizard_form()

    def action_tb_reload_lines(self):
        self.ensure_one()
        self._load_tb_override_lines(preserve_overrides=True)
        return self._reopen_wizard_form()

    def action_tb_reset_overrides(self):
        self.ensure_one()
        period_key = self.env.context.get('tb_period_key')
        target_lines = self.tb_override_line_ids
        if period_key in ('current', 'prior'):
            target_lines = target_lines.filtered(lambda line: line.period_key == period_key)
        for line in target_lines:
            line.write({
                'override_debit': self._to_float(line.system_debit),
                'override_credit': self._to_float(line.system_credit),
                'override_balance': (
                    self._to_float(line.system_debit) - self._to_float(line.system_credit)
                ),
            })
        self._sync_tb_overrides_json()
        try:
            self._get_report_data()
        except Exception as err:
            _logger.debug(
                "Trial balance warning refresh skipped after reset for wizard_id=%s due to: %s",
                self.id,
                err,
            )
        return self._reopen_wizard_form()

    def _get_report_data(self):
        """Get report data for the template"""
        periods = self._get_reporting_periods()
        date_start = periods['date_start']
        date_end = periods['date_end']
        period_category = periods['period_category']
        include_corporate_tax_liability = bool(self.corporate_tax_liability_paid)
        show_prior_year = periods['show_prior_year']
        prior_date_start = periods['prior_date_start']
        prior_date_end = periods['prior_date_end']
        prior_prior_date_start = periods['prior_prior_date_start']
        prior_prior_date_end = periods['prior_prior_date_end']
        balance_sheet_date_start = periods['balance_sheet_date_start']
        prior_balance_sheet_date_start = periods['prior_balance_sheet_date_start']
        prior_prior_balance_sheet_date_start = periods['prior_prior_balance_sheet_date_start']

        def _column_period_word(start_date, end_date):
            if not start_date or not end_date:
                return 'period'
            expected_year_end = start_date + relativedelta(years=1) - relativedelta(days=1)
            return 'year' if expected_year_end == end_date else 'period'

        current_column_period_word = _column_period_word(date_start, date_end)
        prior_column_period_word = _column_period_word(prior_date_start, prior_date_end)
        if not show_prior_year:
            comparative_period_word = 'period'
        elif current_column_period_word == prior_column_period_word:
            comparative_period_word = current_column_period_word
        else:
            comparative_period_word = f"{current_column_period_word} / {prior_column_period_word}"

        fetch_rows_cache = {}
        tb_override_maps = self._build_tb_override_maps()

        current_override_ranges = {
            (balance_sheet_date_start or False, date_end or False),
            (date_start or False, date_end or False),
            (False, date_end or False),
        }
        prior_override_ranges = {
            (prior_balance_sheet_date_start or False, prior_date_end or False),
            (prior_date_start or False, prior_date_end or False),
            (False, prior_date_end or False),
        }

        def _tb_override_period_for_range(range_start, range_end):
            range_key = (range_start or False, range_end or False)
            if range_key in current_override_ranges:
                return 'current'
            if show_prior_year and range_key in prior_override_ranges:
                return 'prior'
            return None

        def _fetch_account_rows(range_start, range_end):
            if not range_end:
                return []
            cache_key = (range_start or False, range_end or False)
            cached_rows = fetch_rows_cache.get(cache_key)
            if cached_rows is not None:
                return [dict(row) for row in cached_rows]

            rows = self._fetch_grouped_account_rows(range_start, range_end)
            period_key = _tb_override_period_for_range(range_start, range_end)
            if period_key:
                rows = self._apply_tb_overrides_to_rows(
                    rows,
                    tb_override_maps.get(period_key) or {},
                )
            cached_payload = [dict(row) for row in rows]
            fetch_rows_cache[cache_key] = cached_payload
            return [dict(row) for row in cached_payload]

        def _build_prefix_totals(rows):
            prefix_lengths = (1, 2, 4, 6, 8)
            totals = {length: {} for length in prefix_lengths}
            for row in rows:
                code = row.get('code') or ''
                if not code:
                    continue
                balance = row.get('balance', 0.0)
                code_length = len(code)
                for length in prefix_lengths:
                    if code_length < length:
                        continue
                    key = code[:length]
                    bucket = totals[length]
                    bucket[key] = bucket.get(key, 0.0) + balance
            return totals

        def _sum_exact_account_totals(prefix_totals, account_codes):
            account_totals = prefix_totals.get(8, {})
            return sum(account_totals.get(code, 0.0) for code in account_codes)

        def _corporate_tax_liability_total(prefix_totals):
            if not include_corporate_tax_liability:
                return 0.0
            return _sum_exact_account_totals(prefix_totals, CT_LIABILITY_ACCOUNT_CODES)


        def _build_section_totals(group_totals, mapping):
            totals = {}
            for label, codes in mapping.items():
                total = 0.0
                for code in codes:
                    total += group_totals.get(code, 0.0)
                totals[label] = total
            return totals

        def _build_group_label_map(mappings):
            labels = {}
            for mapping in mappings:
                for label, codes in mapping.items():
                    for code in codes:
                        labels[code] = label
            return labels

        def _annotate_rows(rows, group_labels):
            annotated = []
            for row in rows:
                code = row.get('code', '')
                group_code = code[:4] if len(code) >= 4 else ''
                data = dict(row)
                data['group_code'] = group_code
                data['sub_group_code'] = code[:6] if len(code) >= 6 else ''
                data['two_digit_code'] = code[:2] if len(code) >= 2 else ''
                data['one_digit_code'] = code[:1] if len(code) >= 1 else ''
                data['section'] = group_labels.get(group_code, 'Other')
                annotated.append(data)
            return annotated

        def _build_subhead_groups(prefix_totals, subhead_labels):
            grouped = {}
            for code in sorted(prefix_totals):
                name = subhead_labels.get(code)
                if not name:
                    continue
                group_code = code[:4]
                grouped.setdefault(group_code, []).append({
                    'code': code,
                    'name': name,
                    'balance': prefix_totals[code],
                })
            return grouped

        def _build_account_head_map(rows):
            accounts = {}
            for row in rows:
                code = row.get('code') or ''
                if len(code) < 8:
                    continue
                group_code = code[:4]
                accounts.setdefault(group_code, {})
                accounts[group_code][code] = {
                    'code': code,
                    'name': row.get('name') or '',
                    'balance': row.get('balance', 0.0),
                }
            return accounts

        def _account_head_balance(account_heads, account_code):
            group_code = account_code[:4]
            return account_heads.get(group_code, {}).get(account_code, {}).get('balance', 0.0)

        def _equity_head_balance(account_heads, account_code):
            # Equity accounts are credit-balanced in Odoo (negative); invert for display.
            return -_account_head_balance(account_heads, account_code)

        owner_current_account_code = '12040101'
        owner_current_account_group_code = owner_current_account_code[:4]
        equity_group_code = '3101'

        def _owner_account_balance(account_heads):
            return (
                account_heads.get(owner_current_account_group_code, {})
                .get(owner_current_account_code, {})
                .get('balance', 0.0)
            )

        def _reclassify_owner_account_group_total(group_totals, account_heads):
            adjusted = dict(group_totals)
            owner_balance = _owner_account_balance(account_heads)
            if not owner_balance:
                return adjusted
            adjusted[owner_current_account_group_code] = (
                adjusted.get(owner_current_account_group_code, 0.0) - owner_balance
            )
            adjusted[equity_group_code] = adjusted.get(equity_group_code, 0.0) + owner_balance
            return adjusted

        def _collect_note_lines(group_codes, current_map, prev_map):
            if len(group_codes) == 1 and group_codes[0] == '1204':
                cash_prefix = '120401'
                bank_prefix = '120402'

                def _sum_prefix(prefix, account_map):
                    total = 0.0
                    for code, info in account_map.get('1204', {}).items():
                        if code == owner_current_account_code:
                            continue
                        if code.startswith(prefix):
                            total += info.get('balance', 0.0)
                    return total

                current_bank = _sum_prefix(bank_prefix, current_map)
                prev_bank = _sum_prefix(bank_prefix, prev_map)
                current_cash = _sum_prefix(cash_prefix, current_map)
                prev_cash = _sum_prefix(cash_prefix, prev_map)

                lines = []
                if current_bank or prev_bank:
                    lines.append({
                        'code': bank_prefix,
                        'name': 'Cash in bank',
                        'current': current_bank,
                        'prev': prev_bank,
                    })
                if current_cash or prev_cash:
                    lines.append({
                        'code': cash_prefix,
                        'name': 'Cash in hand',
                        'current': current_cash,
                        'prev': prev_cash,
                    })
                if lines:
                    return lines

            account_codes = set()
            for group_code in group_codes:
                account_codes.update(current_map.get(group_code, {}).keys())
                account_codes.update(prev_map.get(group_code, {}).keys())

            lines = []
            for account_code in sorted(account_codes):
                group_code = account_code[:4]
                current_info = current_map.get(group_code, {}).get(account_code, {})
                prev_info = prev_map.get(group_code, {}).get(account_code, {})
                current_balance = current_info.get('balance', 0.0)
                prev_balance = prev_info.get('balance', 0.0)
                if not (current_balance or prev_balance):
                    continue
                lines.append({
                    'code': account_code,
                    'name': current_info.get('name') or prev_info.get('name') or '',
                    'current': current_balance,
                    'prev': prev_balance,
                })

            if len(group_codes) == 1 and group_codes[0] == equity_group_code:
                current_owner_info = current_map.get(owner_current_account_group_code, {}).get(
                    owner_current_account_code, {}
                )
                prev_owner_info = prev_map.get(owner_current_account_group_code, {}).get(
                    owner_current_account_code, {}
                )
                current_owner_balance = current_owner_info.get('balance', 0.0)
                prev_owner_balance = prev_owner_info.get('balance', 0.0)
                if current_owner_balance or prev_owner_balance:
                    lines.append({
                        'code': owner_current_account_code,
                        'name': (
                            current_owner_info.get('name')
                            or prev_owner_info.get('name')
                            or 'Owner current account'
                        ),
                        'current': current_owner_balance,
                        'prev': prev_owner_balance,
                    })
                    lines.sort(key=lambda line: line.get('code') or '')
            return lines

        def _label_group_totals(group_totals, labels):
            totals = {}
            for code, label in labels.items():
                totals[label] = totals.get(label, 0.0) + group_totals.get(code, 0.0)
            return totals

        NON_CURRENT_ASSET_GROUPS = {
            'Property and equipment': ['1101'],
            'Right of use of assets': ['1102'],
            'Capital work in progress': ['1103'],
            'Long term investment': ['1104', '1108'],
            'Long term deposits': ['1105'],
            'Long term loans receivable': ['1106'],
            'Intangible assets': ['1107'],
            'Deferred tax': ['1109'],
        }
        CURRENT_ASSET_GROUPS = {
            'Inventory': ['1201'],
            'Accounts receivable': ['1202'],
            'Prepayment, deposits, advances and other receivables': ['1203'],
            'Stripe and other wallets': ['1205'],
            'Cash and bank balances': ['1204'],
            'Interbank transfer': ['1206'],
        }
        NON_CURRENT_LIABILITY_GROUPS = {
            'Long term loan': ['2101'],
            'Long term loans payable - related parties': ['2102'],
            'Deferred tax': ['2103'],
            'Employee benefits': ['2104'],
        }
        CURRENT_LIABILITY_GROUPS = {
            'Short term borrowings': ['2201'],
            'Trade payables': ['2202'],
            'Other payables': ['2203'],
            'Corporate tax liability': ['2204'],
        }
        EQUITY_GROUPS = {
            'Equity': ['3101'],
        }
        REVENUE_GROUPS = {
            'Revenue': ['4101', '4102'],
        }
        OTHER_INCOME_GROUPS = {
            'Other income': ['4103'],
        }
        COST_OF_REVENUE_GROUPS = {
            'Cost of revenue': ['5101'],
        }
        DEPRECIATION_GROUPS = {
            'Depreciation': ['5114'],
        }

        MAIN_HEAD_LABELS = {
            '1101': 'Property Plant And Equipment',
            '1102': 'Right of Use of Assets',
            '1103': 'Capital Work in Progress',
            '1104': 'Long Term Investment',
            '1105': 'Long Term Deposits',
            '1106': 'Long Term Loans Receivable',
            '1107': 'Intangible Assets',
            '1108': 'Long Term Investment',
            '1109': 'Deferred Tax',
            '1201': 'Inventory',
            '1202': 'Accounts Receivable',
            '1203': 'Prepayment, Deposits, Advances & Other Receivables',
            '1204': 'Cash And Bank Balances',
            '1205': 'Stripe & Other Wallets',
            '1206': 'Interbank Transfer',
            '2101': 'Long term loan',
            '2102': 'Long Term Loans Payable - Related Parties',
            '2103': 'Deferred Tax',
            '2104': 'Employee Benefits',
            '2201': 'Short Term Borrowings',
            '2202': 'Trade Payables',
            '2203': 'Other Payables',
            '2204': 'Corporate Tax Liability',
            '3101': 'Equity',
            '4101': 'Revenue',
            '4102': 'Revenue - Related Party',
            '4103': 'Other Income',
            '5101': 'Direct Cost',
            '5102': 'Marketing & Advertisement Expenses',
            '5103': 'IT Expenses',
            '5104': 'Telephone & Internet',
            '5105': 'Professional Fee & Consultancy',
            '5106': 'Legal & Government Fee',
            '5107': "Director's Salary",
            '5108': 'Salaries & Wages',
            '5109': 'Audit & Accounting Expense',
            '5110': 'Rent Expense',
            '5111': 'Utilities',
            '5112': 'Printing & Stationary Expense',
            '5113': 'Travel & Accommodation Expense',
            '5114': 'Depreciation & Amortization',
            '5115': 'Courier Expenses',
            '5116': 'Commission Expense',
            '5117': 'Cleaning Expense',
            '5118': 'Repair & Maintenance',
            '5119': 'Other expenses',
            '5120': 'Penalties & Fines',
            '5121': 'Donation & Charity',
            '5122': 'Other expenses',
            '5123': 'Bank Charges',
            '5124': 'Finance Cost',
            '5125': 'Insurance Expense',
            '5126': 'Fees & Subscriptions',
            '5127': 'Tax expense',
            '5128': 'Office Expense',
        }

        SUBHEAD_LABELS = {
            '110101': 'Land & Buildings',
            '110102': 'Furnitures & Fixtures',
            '110103': 'Leasehold Improvements',
            '110104': 'Vehicles',
            '110105': 'IT Equipments',
            '110106': 'Office Equipments',
            '110201': 'ROU Property',
            '110202': 'ROU Vehicles',
            '110203': 'ROU Equipment',
            '110204': 'ROU Property',
            '110205': 'ROU Vehicles',
            '110206': 'ROU Equipment',
            '110301': 'Capital Work in Progress',
            '110401': 'Investments in funds',
            '110501': 'Long Term Deposits',
            '110601': 'Related Parties',
            '110701': 'Software & Licences',
            '110702': 'Website & Blogs',
            '110703': 'Other Intangible Assets',
            '110801': 'Group Investments',
            '110901': 'Deferred Tax',
            '120101': 'Inventory',
            '120201': 'Trade Receivables',
            '120202': 'Retention Receivables',
            '120301': 'Advances',
            '120302': 'Prepaid Expenses',
            '120303': 'Other Receivables',
            '120304': 'VAT Receivable',
            '120401': 'Cash',
            '120402': 'Bank Accounts',
            '120501': 'Payment Gateways',
            '120502': 'Other Wallet',
            '120601': 'Interbank Transfer',
            '210101': 'Bank Loans',
            '210102': 'Related Party Loan',
            '210103': 'Lease Liabilities',
            '210201': 'Long Term Loans Payable - Related Parties',
            '210301': 'Deferred Tax',
            '210401': 'End of Service Benefits',
            '220101': 'Bank Overdraft',
            '220102': 'Short term loan',
            '220103': 'Credit Card Payable',
            '220104': 'Lease Liability',
            '220201': 'Trade Payables',
            '220301': 'Other Payables',
            '220302': 'VAT Payable',
            '220303': 'Accrued Expenses',
            '220401': 'Corporate Tax Liability',
            '310101': 'Share Capital',
            '310102': 'Retained Earnings',
            '310103': 'Reserves',
            '310104': 'Fund Balances',
            '410101': 'Revenue',
            '410201': 'Revenue - Related Party',
            '410301': 'Interest income',
            '410302': 'Dividend income',
            '410303': 'Rental Income',
            '410304': 'Miscellaneous Income',
            '410305': 'Foreign Exchange Gain',
            '410306': 'Fair Value Gains',
            '410307': 'Endowment Income',
            '410308': 'Donations',
            '410309': 'Grants',
            '410310': 'Write Back',
            '510101': 'Cost of Sales',
            '510102': 'Cost of Goods',
            '510103': 'Cost of Sales',
            '510104': 'Event Fee',
            '510201': 'Marketing & Advertisement Expenses',
            '510301': 'IT & Software Expenses',
            '510401': 'Telephone & Internet',
            '510501': 'Professional Fee & Consultancy',
            '510601': 'Legal & Government Fee',
            '510701': "Director's Salary",
            '510801': 'Office Staff Salaries',
            '510802': 'Coaching Staff Salaries',
            '510803': 'Employee Benefits & Allowances',
            '510804': 'Bonus & Incentives',
            '510805': 'Staff Welfare',
            '510901': 'Audit Fee',
            '510902': 'Accounting & Bookkeeping Fee',
            '511001': 'Office Rent',
            '511002': 'Miscellaneous Rent',
            '511101': 'Office Utilities',
            '511201': 'Printing Cost',
            '511202': 'Stationary',
            '511301': 'Travel & Accommodation Expense',
            '511302': 'Local Accommodation Expense',
            '511401': 'Depreciation Expense',
            '511402': 'Amortization Expense',
            '511501': 'Courier Expenses',
            '511601': 'Commission Expense',
            '511701': 'Cleaning Expense',
            '511801': 'Repair & Maintenance',
            '511901': 'Other Expenses',
            '512001': 'Penalties & Fines',
            '512101': 'Donation & Charity',
            '512201': 'Other Expenses',
            '512202': 'Charitable Activities',
            '512301': 'Bank Charges & Fees',
            '512401': 'Interest Expense',
            '512501': 'Business Insurance',
            '512601': 'Trade License',
            '512602': 'Establishment Card',
            '512603': 'Visa Fee',
            '512701': 'Corporate Tax Expense',
            '512801': 'Office Expense',
        }

        group_labels = _build_group_label_map([
            NON_CURRENT_ASSET_GROUPS,
            CURRENT_ASSET_GROUPS,
            NON_CURRENT_LIABILITY_GROUPS,
            CURRENT_LIABILITY_GROUPS,
            EQUITY_GROUPS,
            REVENUE_GROUPS,
            OTHER_INCOME_GROUPS,
            COST_OF_REVENUE_GROUPS,
            DEPRECIATION_GROUPS,
        ])

        # Current year balances (as of date_end)
        current_rows = _fetch_account_rows(balance_sheet_date_start, date_end)
        current_account_heads = _build_account_head_map(current_rows)
        if not include_corporate_tax_liability:
            current_account_heads = dict(current_account_heads)
            current_account_heads.pop('2204', None)
        current_prefix_totals = _build_prefix_totals(current_rows)
        current_group_totals = dict(
            _reclassify_owner_account_group_total(current_prefix_totals[4], current_account_heads)
        )
        current_group_totals['2204'] = _corporate_tax_liability_total(current_prefix_totals)
        current_main_heads = _label_group_totals(current_group_totals, MAIN_HEAD_LABELS)
        current_subheads_by_group = _build_subhead_groups(
            current_prefix_totals[6],
            SUBHEAD_LABELS,
        )
        current_ppe_subheads = {
            row['name']: row['balance']
            for row in current_subheads_by_group.get('1101', [])
        }
        current_ppe_total = sum(current_ppe_subheads.values())

        annotated_rows = _annotate_rows(current_rows, group_labels)
        balance_rows = [row for row in annotated_rows if row['one_digit_code'] in ('1', '2', '3')]
        period_rows = _fetch_account_rows(date_start, date_end)
        period_account_heads = _build_account_head_map(period_rows)
        period_prefix_totals = _build_prefix_totals(period_rows)
        period_group_totals = period_prefix_totals[4]
        period_annotated_rows = _annotate_rows(period_rows, group_labels)
        profit_loss_accounts = [
            row for row in period_annotated_rows
            if row['one_digit_code'] in ('4', '5')
        ]



        current_assets = _build_section_totals(current_group_totals, CURRENT_ASSET_GROUPS)
        non_current_assets = _build_section_totals(current_group_totals, NON_CURRENT_ASSET_GROUPS)
        current_liabilities = _build_section_totals(current_group_totals, CURRENT_LIABILITY_GROUPS)
        non_current_liabilities = _build_section_totals(current_group_totals, NON_CURRENT_LIABILITY_GROUPS)
        equity = _build_section_totals(current_group_totals, EQUITY_GROUPS)

        current_assets_total = sum(current_assets.values())
        non_current_assets_total = sum(non_current_assets.values())
        total_assets = current_assets_total + non_current_assets_total


        current_liabilities_total = sum(current_liabilities.values())
        non_current_liabilities_total = sum(non_current_liabilities.values())
        total_liabilities = -(current_liabilities_total + non_current_liabilities_total)

        total_equity = sum(equity.values())
        total_of_equity_and_liabilities = total_equity + total_liabilities

        # PnL (current year)
        revenue_total = sum(
            period_group_totals.get(code, 0.0)
            for code in REVENUE_GROUPS['Revenue']
        )
        revenue_total = -(revenue_total)
        other_income_total = sum(
            period_group_totals.get(code, 0.0)
            for code in OTHER_INCOME_GROUPS['Other income']
        )
        other_income_total = -(other_income_total)
        cost_of_revenue_total = sum(
            period_group_totals.get(code, 0.0)
            for code in COST_OF_REVENUE_GROUPS['Cost of revenue']
        )
        depreciation_total = sum(
            period_group_totals.get(code, 0.0)
            for code in DEPRECIATION_GROUPS['Depreciation']
        )
        expense_total = period_prefix_totals[2].get('51', 0.0)
        tax_expense_ledger = _sum_exact_account_totals(period_prefix_totals, CT_EXPENSE_ACCOUNT_CODES)
        operating_expenses_total = expense_total - cost_of_revenue_total - tax_expense_ledger

        gross_profit = revenue_total - cost_of_revenue_total
        net_profit_before_tax = gross_profit - operating_expenses_total + other_income_total
        tax_amount = tax_expense_ledger
        net_profit_after_tax = net_profit_before_tax - tax_amount

        def _calc_margin(amount, base):
            if not base:
                return 0.0
            return (amount / base) * 100.0

        def _profit_or_loss_text(current, prev=None, use_comparative=False):
            current_is_loss = (current or 0.0) < 0
            if not use_comparative:
                return '(loss)' if current_is_loss else 'profit'

            prev_is_loss = (prev or 0.0) < 0
            if current_is_loss and prev_is_loss:
                return '(loss)'
            if (not current_is_loss) and (not prev_is_loss):
                return 'profit'
            if (not current_is_loss) and prev_is_loss:
                return 'profit / (loss)'
            return '(loss) / profit'

        def _net_label(current, prev, period_word, suffix=None, use_comparative=False):
            label_base = f"Net {_profit_or_loss_text(current, prev, use_comparative)}"
            label = f"{label_base} for the {period_word}"
            if suffix:
                label = f"{label} {suffix}"
            return label

        gross_profit_margin = _calc_margin(gross_profit, revenue_total)
        net_profit_margin = _calc_margin(net_profit_after_tax, revenue_total)

        # Prior year balances (as of prior_date_end)
        prev_rows = _fetch_account_rows(prior_balance_sheet_date_start, prior_date_end)
        prev_account_heads = _build_account_head_map(prev_rows)
        if not include_corporate_tax_liability:
            prev_account_heads = dict(prev_account_heads)
            prev_account_heads.pop('2204', None)
        prev_prefix_totals = _build_prefix_totals(prev_rows)
        prev_group_totals = dict(
            _reclassify_owner_account_group_total(prev_prefix_totals[4], prev_account_heads)
        )
        prev_group_totals['2204'] = _corporate_tax_liability_total(prev_prefix_totals)

        def _sum_prefix_group(prefix_totals, prefixes):
            return sum(prefix_totals.get(code, 0.0) for code in prefixes)

        def _sum_other_receivables(prefix_totals):
            return sum(
                value
                for code, value in prefix_totals.items()
                if code.startswith('1203') and code not in ('120301', '120302')
            )

        def _has_amount(*values):
            return any(abs(value) > 0.0 for value in values)

        advances_current = _sum_prefix_group(current_prefix_totals[6], ['120301'])
        advances_prev = _sum_prefix_group(prev_prefix_totals[6], ['120301'])
        prepaid_current = _sum_prefix_group(current_prefix_totals[6], ['120302'])
        prepaid_prev = _sum_prefix_group(prev_prefix_totals[6], ['120302'])
        other_receivables_current = _sum_other_receivables(current_prefix_totals[6])
        other_receivables_prev = _sum_other_receivables(prev_prefix_totals[6])

        receivable_parts = []
        if _has_amount(prepaid_current, prepaid_prev):
            receivable_parts.append('Prepayment')
        if _has_amount(advances_current, advances_prev):
            receivable_parts.append('Advances')
        if _has_amount(other_receivables_current, other_receivables_prev):
            receivable_parts.append('Other receivables')

        receivable_label = None
        if receivable_parts:
            if len(receivable_parts) == 1:
                receivable_label = receivable_parts[0]
            elif len(receivable_parts) == 2:
                receivable_label = f"{receivable_parts[0]} and {receivable_parts[1]}"
            else:
                receivable_label = f"{receivable_parts[0]}, {receivable_parts[1]} and {receivable_parts[2]}"
            MAIN_HEAD_LABELS['1203'] = receivable_label

        prev_prev_rows = _fetch_account_rows(prior_prior_balance_sheet_date_start, prior_prior_date_end)
        prev_prev_account_heads = _build_account_head_map(prev_prev_rows)
        prev_prev_prefix_totals = _build_prefix_totals(prev_prev_rows)
        prev_prev_group_totals = _reclassify_owner_account_group_total(
            prev_prev_prefix_totals[4],
            prev_prev_account_heads,
        )
        prev_period_rows = _fetch_account_rows(prior_date_start, prior_date_end)
        prev_period_account_heads = _build_account_head_map(prev_period_rows)
        prev_period_prefix_totals = _build_prefix_totals(prev_period_rows)
        prev_period_group_totals = prev_period_prefix_totals[4]
        prev_prev_period_rows = _fetch_account_rows(prior_prior_date_start, prior_prior_date_end)
        prev_prev_period_account_heads = _build_account_head_map(prev_prev_period_rows)
        prev_prev_period_prefix_totals = _build_prefix_totals(prev_prev_period_rows)
        prev_prev_period_group_totals = prev_prev_period_prefix_totals[4]
        prev_main_heads = _label_group_totals(prev_group_totals, MAIN_HEAD_LABELS)
        prev_subheads_by_group = _build_subhead_groups(
            prev_prefix_totals[6],
            SUBHEAD_LABELS,
        )
        prev_ppe_subheads = {
            row['name']: row['balance']
            for row in prev_subheads_by_group.get('1101', [])
        }
        prev_ppe_total = sum(prev_ppe_subheads.values())

        prev_current_assets = _build_section_totals(prev_group_totals, CURRENT_ASSET_GROUPS)
        prev_non_current_assets = _build_section_totals(prev_group_totals, NON_CURRENT_ASSET_GROUPS)
        prev_current_liabilities = _build_section_totals(prev_group_totals, CURRENT_LIABILITY_GROUPS)
        prev_non_current_liabilities = _build_section_totals(prev_group_totals, NON_CURRENT_LIABILITY_GROUPS)
        prev_equity = _build_section_totals(prev_group_totals, EQUITY_GROUPS)
        prev_prev_current_assets = _build_section_totals(prev_prev_group_totals, CURRENT_ASSET_GROUPS)

        prev_current_assets_total = sum(prev_current_assets.values())
        prev_non_current_assets_total = sum(prev_non_current_assets.values())
        prev_total_assets = prev_current_assets_total + prev_non_current_assets_total

        prev_current_liabilities_total = sum(prev_current_liabilities.values())
        prev_non_current_liabilities_total = sum(prev_non_current_liabilities.values())
        prev_total_liabilities = -(prev_current_liabilities_total + prev_non_current_liabilities_total)

        prev_equity_total = sum(prev_equity.values())
        prev_total_equity = prev_equity_total
        prev_total_of_equity_and_liabilities = prev_total_equity + prev_total_liabilities

        prev_revenue_total = sum(
            prev_period_group_totals.get(code, 0.0)
            for code in REVENUE_GROUPS['Revenue']
        )
        prev_revenue_total = -(prev_revenue_total)
        prev_other_income_total = sum(
            prev_period_group_totals.get(code, 0.0)
            for code in OTHER_INCOME_GROUPS['Other income']
        )
        prev_other_income_total = -(prev_other_income_total)
        prev_cost_of_revenue_total = sum(
            prev_period_group_totals.get(code, 0.0)
            for code in COST_OF_REVENUE_GROUPS['Cost of revenue']
        )
        prev_depreciation_total = sum(
            prev_period_group_totals.get(code, 0.0)
            for code in DEPRECIATION_GROUPS['Depreciation']
        )
        prev_expense_total = prev_period_prefix_totals[2].get('51', 0.0)
        prev_tax_expense_ledger = _sum_exact_account_totals(
            prev_period_prefix_totals,
            CT_EXPENSE_ACCOUNT_CODES,
        )
        prev_operating_expenses_total = (
            prev_expense_total - prev_cost_of_revenue_total - prev_tax_expense_ledger
        )

        prev_gross_profit = prev_revenue_total - prev_cost_of_revenue_total
        prev_net_profit_before_tax = (
            prev_gross_profit - prev_operating_expenses_total + prev_other_income_total
        )
        prev_tax_amount = prev_tax_expense_ledger
        prev_net_profit_after_tax = prev_net_profit_before_tax - prev_tax_amount

        prev_gross_profit_margin = _calc_margin(prev_gross_profit, prev_revenue_total)
        prev_net_profit_margin = _calc_margin(prev_net_profit_after_tax, prev_revenue_total)

        label_prev_net_profit_before_tax = prev_net_profit_before_tax if show_prior_year else 0.0
        label_prev_net_profit_after_tax = prev_net_profit_after_tax if show_prior_year else 0.0
        period_word = comparative_period_word
        show_after_tax_line = bool(
            (tax_amount or (show_prior_year and prev_tax_amount))
            and (net_profit_after_tax or (show_prior_year and prev_net_profit_after_tax))
        )
        before_tax_suffix = "before tax" if show_after_tax_line else None
        net_profit_before_tax_label = _net_label(
            net_profit_before_tax,
            label_prev_net_profit_before_tax,
            period_word,
            before_tax_suffix,
            show_prior_year,
        )
        net_profit_after_tax_label = _net_label(
            net_profit_after_tax,
            label_prev_net_profit_after_tax,
            period_word,
            "after tax",
            show_prior_year,
        )
        gross_profit_margin_label = (
            f"Gross {_profit_or_loss_text(gross_profit_margin, prev_gross_profit_margin, show_prior_year)} margin"
        )

        prev_prev_revenue_total = sum(
            prev_prev_period_group_totals.get(code, 0.0)
            for code in REVENUE_GROUPS['Revenue']
        )
        prev_prev_revenue_total = -(prev_prev_revenue_total)
        prev_prev_other_income_total = sum(
            prev_prev_period_group_totals.get(code, 0.0)
            for code in OTHER_INCOME_GROUPS['Other income']
        )
        prev_prev_other_income_total = -(prev_prev_other_income_total)
        prev_prev_cost_of_revenue_total = sum(
            prev_prev_period_group_totals.get(code, 0.0)
            for code in COST_OF_REVENUE_GROUPS['Cost of revenue']
        )
        prev_prev_depreciation_total = sum(
            prev_prev_period_group_totals.get(code, 0.0)
            for code in DEPRECIATION_GROUPS['Depreciation']
        )
        prev_prev_expense_total = prev_prev_period_prefix_totals[2].get('51', 0.0)
        prev_prev_tax_expense_ledger = _sum_exact_account_totals(
            prev_prev_period_prefix_totals,
            CT_EXPENSE_ACCOUNT_CODES,
        )
        prev_prev_operating_expenses_total = (
            prev_prev_expense_total - prev_prev_cost_of_revenue_total - prev_prev_tax_expense_ledger
        )
        prev_prev_gross_profit = prev_prev_revenue_total - prev_prev_cost_of_revenue_total
        prev_prev_net_profit_before_tax = (
            prev_prev_gross_profit
            - prev_prev_operating_expenses_total
            + prev_prev_other_income_total
        )
        prev_prev_tax_amount = prev_prev_tax_expense_ledger
        prev_prev_net_profit_after_tax = prev_prev_net_profit_before_tax - prev_prev_tax_amount

        note_number = 5
        note_numbers = {}
        note_sections = []
        show_shareholder_note = bool(self.show_shareholder_note)
        note_labels = dict(MAIN_HEAD_LABELS)
        if receivable_label:
            note_labels['1203'] = receivable_label

        note_prev_account_heads = prev_account_heads if show_prior_year else {}
        note_prev_group_totals = prev_group_totals if show_prior_year else {}
        note_prev_period_account_heads = prev_period_account_heads if show_prior_year else {}
        note_prev_period_group_totals = prev_period_group_totals if show_prior_year else {}

        def _sort_operating_expense_lines(lines):
            top_codes = ['51070101', '51080101']
            bottom_codes = ['51230101', '51220102']
            top_set = set(top_codes)
            bottom_set = set(bottom_codes)
            top_lines = [line for code in top_codes for line in lines if line.get('code') == code]
            bottom_lines = [line for code in bottom_codes for line in lines if line.get('code') == code]
            middle_lines = [
                line for line in lines
                if line.get('code') not in top_set and line.get('code') not in bottom_set
            ]
            middle_lines.sort(key=lambda line: (line.get('name') or line.get('code') or '').lower())
            return top_lines + middle_lines + bottom_lines

        def _aggregate_operating_expense_lines(lines):
            fee_codes = {'51260101', '51260201', '51260301'}
            fee_current = 0.0
            fee_prev = 0.0
            aggregated = []
            for line in lines:
                if line.get('code') in fee_codes:
                    fee_current += line.get('current', 0.0)
                    fee_prev += line.get('prev', 0.0)
                    continue
                aggregated.append(line)
            if fee_current or fee_prev:
                aggregated.append({
                    'code': 'fee_subscription',
                    'name': 'Fee and subscription',
                    'current': fee_current,
                    'prev': fee_prev,
                })
            return aggregated

        def _add_note(key, label, group_codes, current_total, prev_total, current_map=None, prev_map=None):
            nonlocal note_number
            lines = _collect_note_lines(
                group_codes,
                current_map or current_account_heads,
                prev_map or note_prev_account_heads,
            )
            if not lines:
                return
            if key == 'pl_revenue':
                # Revenue note should be presented as a single summarized line.
                lines = [{
                    'code': 'pl_revenue_total',
                    'name': 'Revenue',
                    'current': current_total,
                    'prev': prev_total,
                }]
            if key == 'pl_opex':
                for line in lines:
                    line_name = (line.get('name') or '').strip().lower()
                    if line_name == 'loss on difference on exchange':
                        line['name'] = 'Exchange loss'
                lines = _aggregate_operating_expense_lines(lines)
                lines = _sort_operating_expense_lines(lines)
            if key == 'pl_cost':
                for line in lines:
                    line_name = (line.get('name') or '').strip().lower()
                    if line_name in ('stock purchase', 'stock purchases'):
                        line['name'] = 'Purchase'
            note_numbers[key] = note_number
            note_sections.append({
                'number': note_number,
                'label': label,
                'lines': lines,
                'total_current': current_total,
                'total_prev': prev_total,
            })
            note_number += 1

        def _ensure_retained_earnings_note():
            nonlocal note_number
            if 'retained_earnings' in note_numbers:
                return
            note_numbers['retained_earnings'] = note_number
            note_sections.append({
                'number': note_number,
                'label': 'Retained earnings',
                'lines': [],
                'total_current': 0.0,
                'total_prev': 0.0,
            })
            note_number += 1

        def _add_equity_account_note(key, label, account_code):
            nonlocal note_number
            if key in note_numbers:
                return
            current_total = _equity_head_balance(current_account_heads, account_code)
            prev_total = _equity_head_balance(note_prev_account_heads, account_code)
            if not (current_total or prev_total):
                return
            group_code = account_code[:4]
            current_info = current_account_heads.get(group_code, {}).get(account_code, {})
            prev_info = note_prev_account_heads.get(group_code, {}).get(account_code, {})
            note_numbers[key] = note_number
            note_sections.append({
                'number': note_number,
                'label': label,
                'lines': [{
                    'code': account_code,
                    'name': current_info.get('name') or prev_info.get('name') or label,
                    'current': current_total,
                    'prev': prev_total,
                }],
                'total_current': current_total,
                'total_prev': prev_total,
            })
            note_number += 1

        def _ensure_equity_detail_notes():
            _ensure_retained_earnings_note()
            note_numbers['statutory_reserves_equity'] = note_numbers.get('retained_earnings')

        non_current_codes = ['1101', '1102', '1103', '1104', '1105', '1106', '1107', '1108', '1109']
        current_codes = ['1201', '1202', '1203', '1205', '1204', '1206']
        equity_codes = ['3101']
        non_current_liability_codes = ['2101', '2102', '2103', '2104']
        current_liability_codes = ['2201', '2202', '2203', '2204']
        aggregated_payables_codes = set(current_liability_codes)
        for code in non_current_codes + current_codes + equity_codes + non_current_liability_codes + current_liability_codes:
            if code in aggregated_payables_codes:
                continue
            if code == '3101' and not show_shareholder_note:
                _ensure_equity_detail_notes()
                continue
            current_total = current_group_totals.get(code, 0.0)
            prev_total = note_prev_group_totals.get(code, 0.0)
            if not (current_total or prev_total):
                if code == '3101':
                    if show_shareholder_note and code not in note_numbers:
                        note_numbers[code] = note_number
                        note_sections.append({
                            'number': note_number,
                            'label': note_labels.get(code, code),
                            'lines': [],
                            'total_current': 0.0,
                            'total_prev': 0.0,
                        })
                        note_number += 1
                    _ensure_equity_detail_notes()
                continue
            _add_note(code, note_labels.get(code, code), [code], current_total, prev_total)
            if code == '3101':
                _ensure_equity_detail_notes()

        aggregated_payables_current = sum(
            current_group_totals.get(code, 0.0)
            for code in aggregated_payables_codes
        )
        aggregated_payables_prev = sum(
            note_prev_group_totals.get(code, 0.0)
            for code in aggregated_payables_codes
        )
        if aggregated_payables_current or aggregated_payables_prev:
            _add_note(
                'other_payables',
                'Other payables',
                sorted(aggregated_payables_codes),
                aggregated_payables_current,
                aggregated_payables_prev,
            )

        expense_group_codes = sorted({
            code for code in current_group_totals
            if code.startswith('51')
        } | {
            code for code in note_prev_group_totals
            if code.startswith('51')
        })
        cost_codes = set(COST_OF_REVENUE_GROUPS['Cost of revenue'])
        depreciation_codes = set(DEPRECIATION_GROUPS['Depreciation'])
        tax_codes = {'5127'}
        operating_expense_codes = [
            code
            for code in expense_group_codes
            if code not in cost_codes | depreciation_codes | tax_codes
        ]

        note_prev_revenue_total = prev_revenue_total if show_prior_year else 0.0
        note_prev_cost_of_revenue_total = prev_cost_of_revenue_total if show_prior_year else 0.0
        note_prev_operating_expenses_total = prev_operating_expenses_total if show_prior_year else 0.0
        if revenue_total or note_prev_revenue_total:
            _add_note(
                'pl_revenue',
                'Revenue',
                REVENUE_GROUPS['Revenue'],
                revenue_total,
                note_prev_revenue_total,
                current_map=period_account_heads,
                prev_map=note_prev_period_account_heads,
            )
        if cost_of_revenue_total or note_prev_cost_of_revenue_total:
            _add_note(
                'pl_cost',
                'Direct cost',
                COST_OF_REVENUE_GROUPS['Cost of revenue'],
                cost_of_revenue_total,
                note_prev_cost_of_revenue_total,
                current_map=period_account_heads,
                prev_map=note_prev_period_account_heads,
            )
        if operating_expenses_total or note_prev_operating_expenses_total:
            current_total = sum(period_group_totals.get(code, 0.0) for code in operating_expense_codes)
            prev_total = sum(note_prev_period_group_totals.get(code, 0.0) for code in operating_expense_codes)
            _add_note(
                'pl_opex',
                'Operating expenses',
                operating_expense_codes,
                current_total,
                prev_total,
                current_map=period_account_heads,
                prev_map=note_prev_period_account_heads,
            )

        post_table_note_keys = []
        if self.show_related_parties_note:
            post_table_note_keys.append('related_parties')
        post_table_note_keys.extend([
            'risk_management',
            'financial_assets_liabilities',
            'contingent_liabilities',
            'general_notes',
        ])
        for key in post_table_note_keys:
            note_numbers[key] = note_number
            note_number += 1
        last_note_number = note_number - 1

        opening_date_end = date_start - relativedelta(days=1) if date_start else False
        opening_rows = _fetch_account_rows(False, opening_date_end)
        opening_prefix_totals = _build_prefix_totals(opening_rows)
        opening_account_heads = _build_account_head_map(opening_rows)
        if not include_corporate_tax_liability:
            opening_account_heads = dict(opening_account_heads)
            opening_account_heads.pop('2204', None)
        opening_group_totals = dict(
            _reclassify_owner_account_group_total(opening_prefix_totals[4], opening_account_heads)
        )
        opening_group_totals['2204'] = _corporate_tax_liability_total(opening_prefix_totals)

        if show_prior_year:
            prior_opening_date_end = prior_date_start - relativedelta(days=1) if prior_date_start else False
            prior_opening_rows = _fetch_account_rows(False, prior_opening_date_end)
            prior_opening_prefix_totals = _build_prefix_totals(prior_opening_rows)
            prior_opening_account_heads = _build_account_head_map(prior_opening_rows)
            if not include_corporate_tax_liability:
                prior_opening_account_heads = dict(prior_opening_account_heads)
                prior_opening_account_heads.pop('2204', None)
        else:
            prior_opening_prefix_totals = _build_prefix_totals([])
            prior_opening_account_heads = {}
        prior_opening_group_totals = _reclassify_owner_account_group_total(
            prior_opening_prefix_totals[4],
            prior_opening_account_heads,
        )
        prior_opening_group_totals = dict(prior_opening_group_totals)
        prior_opening_group_totals['2204'] = _corporate_tax_liability_total(prior_opening_prefix_totals)
        soce_prior_opening_label_date = self.soce_prior_opening_label_date or prior_date_start

        ###### Variables for Cash Flow Statement ######
        eosb_liability_code = '21040101'
        dividend_paid_code = '31010202'

        def _liability_to_positive(value):
            return -value

        current_depreciation_total = period_group_totals.get('5114', 0.0)
        prior_depreciation_total = prev_period_group_totals.get('5114', 0.0)

        eosb_movement = _liability_to_positive(period_prefix_totals[8].get(eosb_liability_code, 0.0))
        prior_eosb_movement = _liability_to_positive(prev_period_prefix_totals[8].get(eosb_liability_code, 0.0))

        end_service_benefits_adjustment = max(eosb_movement, 0.0)
        prior_end_service_benefits_adjustment = max(prior_eosb_movement, 0.0)
        end_service_benefits_paid = -max(-eosb_movement, 0.0)
        prior_end_service_benefits_paid = -max(-prior_eosb_movement, 0.0)

        operating_cashflow_before_working_capital = (
            net_profit_before_tax
            + current_depreciation_total
            + end_service_benefits_adjustment
        )
        prior_operating_cashflow_before_working_capital = (
            prev_net_profit_before_tax
            + prior_depreciation_total
            + prior_end_service_benefits_adjustment
        )

        current_wc_assets = (
            current_group_totals.get('1201', 0.0)
            + current_group_totals.get('1202', 0.0)
            + current_group_totals.get('1203', 0.0)
            + current_group_totals.get('1205', 0.0)
        )
        opening_wc_assets = (
            opening_group_totals.get('1201', 0.0)
            + opening_group_totals.get('1202', 0.0)
            + opening_group_totals.get('1203', 0.0)
            + opening_group_totals.get('1205', 0.0)
        )
        prior_wc_assets = (
            prev_group_totals.get('1201', 0.0)
            + prev_group_totals.get('1202', 0.0)
            + prev_group_totals.get('1203', 0.0)
            + prev_group_totals.get('1205', 0.0)
        )
        prior_opening_wc_assets = (
            prior_opening_group_totals.get('1201', 0.0)
            + prior_opening_group_totals.get('1202', 0.0)
            + prior_opening_group_totals.get('1203', 0.0)
            + prior_opening_group_totals.get('1205', 0.0)
        )

        change_in_current_assets = -(current_wc_assets - opening_wc_assets)
        prior_change_in_current_assets = -(prior_wc_assets - prior_opening_wc_assets)

        current_wc_liabilities = _liability_to_positive(
            current_group_totals.get('2202', 0.0) + current_group_totals.get('2203', 0.0)
        )
        opening_wc_liabilities = _liability_to_positive(
            opening_group_totals.get('2202', 0.0) + opening_group_totals.get('2203', 0.0)
        )
        prior_wc_liabilities = _liability_to_positive(
            prev_group_totals.get('2202', 0.0) + prev_group_totals.get('2203', 0.0)
        )
        prior_opening_wc_liabilities = _liability_to_positive(
            prior_opening_group_totals.get('2202', 0.0) + prior_opening_group_totals.get('2203', 0.0)
        )

        change_in_current_liabilities = current_wc_liabilities - opening_wc_liabilities
        prior_change_in_current_liabilities = prior_wc_liabilities - prior_opening_wc_liabilities

        if include_corporate_tax_liability:
            current_tax_liability = _liability_to_positive(
                _corporate_tax_liability_total(current_prefix_totals)
            )
            opening_tax_liability = _liability_to_positive(
                _corporate_tax_liability_total(opening_prefix_totals)
            )
            prior_tax_liability = _liability_to_positive(
                _corporate_tax_liability_total(prev_prefix_totals)
            )
            prior_opening_tax_liability = _liability_to_positive(
                _corporate_tax_liability_total(prior_opening_prefix_totals)
            )
            current_tax_expense = _sum_exact_account_totals(period_prefix_totals, CT_EXPENSE_ACCOUNT_CODES)
            prior_tax_expense = _sum_exact_account_totals(prev_period_prefix_totals, CT_EXPENSE_ACCOUNT_CODES)

            corporate_tax_paid = -(opening_tax_liability + current_tax_expense - current_tax_liability)
            prior_corporate_tax_paid = -(prior_opening_tax_liability + prior_tax_expense - prior_tax_liability)
        else:
            corporate_tax_paid = 0.0
            prior_corporate_tax_paid = 0.0

        net_cash_generated_from_operations = (
            operating_cashflow_before_working_capital
            + change_in_current_assets
            + change_in_current_liabilities
            + corporate_tax_paid
            + end_service_benefits_paid
        )
        prior_net_cash_generated_from_operations = (
            prior_operating_cashflow_before_working_capital
            + prior_change_in_current_assets
            + prior_change_in_current_liabilities
            + prior_corporate_tax_paid
            + prior_end_service_benefits_paid
        )

        current_property = -(period_group_totals.get('1101', 0.0))
        prior_property = -(prev_period_group_totals.get('1101', 0.0))
        net_cash_generated_from_investing_activities = current_property
        prior_net_cash_generated_from_investing_activities = prior_property

        company = self.company_id
        share_capital_account_code = '31010101'
        statement_share_capital_total = _equity_head_balance(current_account_heads, share_capital_account_code)
        statement_prev_share_capital_total = _equity_head_balance(prev_account_heads, share_capital_account_code)
        statement_prev_prev_share_capital_total = _equity_head_balance(
            prev_prev_account_heads,
            share_capital_account_code,
        )
        statement_prior_opening_share_capital_total = _equity_head_balance(
            prior_opening_account_heads,
            share_capital_account_code,
        )

        # Financial statement tables: pick share capital from GL account 31010101.
        if show_prior_year:
            # 2-year templates: show movement in paid-up capital instead of closing balance.
            paid_up_capital = statement_share_capital_total - statement_prev_share_capital_total
            prior_paid_up_capital = (
                statement_prev_share_capital_total - statement_prior_opening_share_capital_total
            )
        else:
            paid_up_capital = statement_share_capital_total
            prior_paid_up_capital = 0.0

        current_dividend = current_prefix_totals[8].get(dividend_paid_code, 0.0)
        opening_dividend = opening_prefix_totals[8].get(dividend_paid_code, 0.0)
        prior_dividend = prev_prefix_totals[8].get(dividend_paid_code, 0.0)
        prior_opening_dividend = prior_opening_prefix_totals[8].get(dividend_paid_code, 0.0)
        dividend_paid = -(current_dividend - opening_dividend)
        prior_dividend_paid = -(prior_dividend - prior_opening_dividend)

        current_owner_ca = current_prefix_totals[8].get(owner_current_account_code, 0.0)
        opening_owner_ca = opening_prefix_totals[8].get(owner_current_account_code, 0.0)
        prior_owner_ca = prev_prefix_totals[8].get(owner_current_account_code, 0.0)
        prior_opening_owner_ca = prior_opening_prefix_totals[8].get(owner_current_account_code, 0.0)
        owner_current_account = -(current_owner_ca - opening_owner_ca)
        prior_owner_current_account = -(prior_owner_ca - prior_opening_owner_ca)

        net_cash_generated_from_financing_activities = (
            paid_up_capital + dividend_paid + owner_current_account
        )
        prior_net_cash_generated_from_financing_activities = (
            prior_paid_up_capital + prior_dividend_paid + prior_owner_current_account
        )

        current_cash_and_bank = current_group_totals.get('1204', 0.0)
        current_interbank = current_group_totals.get('1206', 0.0)
        opening_cash_and_bank = opening_group_totals.get('1204', 0.0)
        opening_interbank = opening_group_totals.get('1206', 0.0)
        prior_cash_and_bank = prev_group_totals.get('1204', 0.0)
        prior_interbank = prev_group_totals.get('1206', 0.0)
        prior_opening_cash_and_bank = prior_opening_group_totals.get('1204', 0.0)
        prior_opening_interbank = prior_opening_group_totals.get('1206', 0.0)

        current_cash_equivalents = current_cash_and_bank - current_interbank
        opening_cash_equivalents = opening_cash_and_bank - opening_interbank
        prior_cash_equivalents = prior_cash_and_bank - prior_interbank
        prior_opening_cash_equivalents = prior_opening_cash_and_bank - prior_opening_interbank

        net_cash_and_cash_equivalents = (
            net_cash_generated_from_operations
            + net_cash_generated_from_investing_activities
            + net_cash_generated_from_financing_activities
        )
        cash_beginning_year = opening_cash_equivalents
        cash_end_of_year = current_cash_equivalents

        prior_net_cash_and_cash_equivalents = (
            prior_net_cash_generated_from_operations
            + prior_net_cash_generated_from_investing_activities
            + prior_net_cash_generated_from_financing_activities
        )
        prior_cash_beginning_year = prior_opening_cash_equivalents
        prior_cash_end_of_year = prior_cash_equivalents


        ##############################################################################################

        ############# Shareholders Section #############


        # Shareholders START #
        #########################################################################################################
        company = self.company_id
        # Shareholders (explicit variables)
        sh_name_1 = company.shareholder_1
        sh_nat_1 = company.nationality_1
        sh_shares_1 = company.number_of_shares_1
        sh_value_1 = company.share_value_1
        sh_total_1 = company.total_share_1

        sh_name_2 = company.shareholder_2
        sh_nat_2 = company.nationality_2
        sh_shares_2 = company.number_of_shares_2
        sh_value_2 = company.share_value_2
        sh_total_2 = company.total_share_2

        sh_name_3 = company.shareholder_3
        sh_nat_3 = company.nationality_3
        sh_shares_3 = company.number_of_shares_3
        sh_value_3 = company.share_value_3
        sh_total_3 = company.total_share_3

        sh_name_4 = company.shareholder_4
        sh_nat_4 = company.nationality_4
        sh_shares_4 = company.number_of_shares_4
        sh_value_4 = company.share_value_4
        sh_total_4 = company.total_share_4

        sh_name_5 = company.shareholder_5
        sh_nat_5 = company.nationality_5
        sh_shares_5 = company.number_of_shares_5
        sh_value_5 = company.share_value_5
        sh_total_5 = company.total_share_5

        sh_name_6 = company.shareholder_6
        sh_nat_6 = company.nationality_6
        sh_shares_6 = company.number_of_shares_6
        sh_value_6 = company.share_value_6
        sh_total_6 = company.total_share_6

        sh_name_7 = company.shareholder_7
        sh_nat_7 = company.nationality_7
        sh_shares_7 = company.number_of_shares_7
        sh_value_7 = company.share_value_7
        sh_total_7 = company.total_share_7

        sh_name_8 = company.shareholder_8
        sh_nat_8 = company.nationality_8
        sh_shares_8 = company.number_of_shares_8
        sh_value_8 = company.share_value_8
        sh_total_8 = company.total_share_8

        sh_name_9 = company.shareholder_9
        sh_nat_9 = company.nationality_9
        sh_shares_9 = company.number_of_shares_9
        sh_value_9 = company.share_value_9
        sh_total_9 = company.total_share_9

        sh_name_10 = company.shareholder_10
        sh_nat_10 = company.nationality_10
        sh_shares_10 = company.number_of_shares_10
        sh_value_10 = company.share_value_10
        sh_total_10 = company.total_share_10

        share_rows = []
        for i in range(1, 11):
            name = getattr(company, f'shareholder_{i}', False)
            if not name:
                continue
            if not getattr(self, f'owner_include_{i}', False):
                continue
            share_rows.append({
                'name': name,
                'value': getattr(company, f'share_value_{i}', 0.0) or 0.0,
                'shares': getattr(company, f'number_of_shares_{i}', 0) or 0,
                'total': getattr(company, f'total_share_{i}', 0.0) or 0.0,
            })

        authorized_share_capital = sum(row['total'] for row in share_rows)
        total_shares_count = sum(row['shares'] for row in share_rows)
        share_value_default = share_rows[0]['value'] if share_rows else 0.0

        signature_entries = []
        for i in range(1, 11):
            name = getattr(company, f'shareholder_{i}', False)
            if not name or not name.strip():
                continue
            if not getattr(self, f'signature_include_{i}', False):
                continue
            role = (getattr(self, f'signature_role_{i}', 'secondary') or 'secondary').lower()
            is_primary = role == 'primary'
            nationality = getattr(company, f'nationality_{i}', False)
            display_name = f"{name} – {nationality}" if nationality else name
            signature_entries.append({
                'index': i,
                'name': name.strip(),
                'display_name': display_name,
                'is_primary': is_primary,
            })

        if signature_entries and not any(entry['is_primary'] for entry in signature_entries):
            signature_entries[0]['is_primary'] = True

        signature_entries.sort(key=lambda entry: (0 if entry['is_primary'] else 1, entry['index']))

        # Keep one "primary" signatory for existing template rendering.
        for idx, entry in enumerate(signature_entries):
            entry['is_primary'] = idx == 0

        signature_names = [entry['name'] for entry in signature_entries]
        signature_display_names = [entry['display_name'] for entry in signature_entries]

        owner_display_names = []
        for i in range(1, 11):
            name = getattr(company, f'shareholder_{i}', False)
            if not name:
                continue
            if not getattr(self, f'owner_include_{i}', False):
                continue
            nationality = getattr(company, f'nationality_{i}', False)
            if nationality:
                owner_display_names.append(f"{name} – {nationality}")
            else:
                owner_display_names.append(name)

        shareholder_display_names = []
        for i in range(1, 11):
            name = getattr(company, f'shareholder_{i}', False)
            if not name or not name.strip():
                continue
            if not getattr(self, f'director_include_{i}', False):
                continue
            nationality = getattr(company, f'nationality_{i}', False)
            if nationality:
                shareholder_display_names.append(f"{name.strip()} ({nationality})")
            else:
                shareholder_display_names.append(name.strip())

        def _join_names(names):
            if not names:
                return ''
            if len(names) == 1:
                return names[0]
            if len(names) == 2:
                return f"{names[0]} and {names[1]}"
            return f"{', '.join(names[:-1])} and {names[-1]}"

        if not shareholder_display_names:
            fallback_name = (company.shareholder_1 or '').strip()
            fallback_nationality = (company.nationality_1 or '').strip()
            if fallback_name:
                if fallback_nationality:
                    shareholder_display_names = [f"{fallback_name} ({fallback_nationality})"]
                else:
                    shareholder_display_names = [fallback_name]
        shareholder_display_text = _join_names(shareholder_display_names)
        management_director_names = []
        for i in range(1, 11):
            name = getattr(company, f'shareholder_{i}', False)
            if not name or not name.strip():
                continue
            if getattr(self, f'signature_include_{i}', False):
                management_director_names.append(name.strip())
        if not management_director_names:
            fallback_director = (company.shareholder_1 or '').strip()
            if fallback_director:
                management_director_names = [fallback_director]
        management_director_text = _join_names(management_director_names)
        management_director_title = 'Director' if len(management_director_names) <= 1 else 'Directors'
        generated_date = fields.Date.context_today(self)
        generated_date_display = generated_date.strftime("%d/%m/%Y") if generated_date else ''

        def _format_long_date(date_value):
            return date_value.strftime('%d %B %Y') if date_value else ''

        emphasis_note_items = []
        emphasis_note_lines = []
        emphasis_sub_note = 6

        def _next_emphasis_note_ref():
            nonlocal emphasis_sub_note
            note_ref = f'1.{emphasis_sub_note}'
            emphasis_sub_note += 1
            return note_ref

        if self.emphasis_change_period:
            changed_from = _format_long_date(
                fields.Date.to_date(self.emphasis_change_from_date) if self.emphasis_change_from_date else None
            )
            changed_to = _format_long_date(
                fields.Date.to_date(self.emphasis_change_to_date) if self.emphasis_change_to_date else None
            )
            change_period_note_ref = _next_emphasis_note_ref()
            emphasis_note_items.append({
                'key': 'change_period',
                'label': 'Change in Financial Year/Period',
                'note_ref': change_period_note_ref,
                'matter_text': (
                    f"the change in the Entity’s financial reporting period from {changed_from} "
                    f"to {changed_to}"
                ),
            })
            emphasis_note_lines.append({
                'note_ref': change_period_note_ref,
                'note_text': (
                    f"The Entity previously issued audited financial statements for the period ended {changed_from}. "
                    f"To align with regulatory requirements, the Entity has revised its financial reporting period to "
                    f"{changed_to}, and this change has been applied retrospectively. Accordingly, the comparative "
                    f"information presented is for the period ended {changed_to}."
                ),
            })

        if self.additional_note_legal_status_change:
            legal_change_date = _format_long_date(
                fields.Date.to_date(self.emphasis_legal_change_date) if self.emphasis_legal_change_date else None
            )
            legal_status_from = (self.emphasis_legal_status_from or '').strip()
            legal_status_to = (self.emphasis_legal_status_to or '').strip()
            legal_registration_authority = (getattr(company, 'free_zone', '') or '').strip()
            if not legal_registration_authority:
                legal_registration_authority = 'Department of Economy and Tourism'
            emphasis_note_lines.append({
                'note_ref': _next_emphasis_note_ref(),
                'note_text': (
                    f'On {legal_change_date}, the Entity officially changed its legal status from "{legal_status_from}" '
                    f'to "{legal_status_to}". The change was approved by the Director and has been '
                    f'registered with the {legal_registration_authority}, Dubai, United Arab Emirates.'
                ),
            })

        if self.additional_note_no_active_bank_account:
            emphasis_note_lines.append({
                'note_ref': _next_emphasis_note_ref(),
                'note_text': (
                    f"{(self.company_name or 'The Entity').upper()} (the Entity) currently has no active bank account. "
                    'All expenses were paid by the Director from his personal resources.'
                ),
            })

        if self.additional_note_business_bank_account_opened:
            emphasis_note_lines.append({
                'note_ref': _next_emphasis_note_ref(),
                'note_text': (
                    f"{self.company_name or 'The Entity'} (the Entity) initially conducted transactions using a "
                    'personal account instead of maintaining a dedicated business bank account. Subsequently, '
                    'post year-end, a business bank account has been opened and is now fully operational.'
                ),
            })

        if self.emphasis_correction_error:
            correction_note_ref = _next_emphasis_note_ref()
            emphasis_note_items.append({
                'key': 'correction_error',
                'label': 'Correction of Error',
                'note_ref': correction_note_ref,
                'matter_text': "the correction of a prior period error",
            })
            emphasis_note_lines.append({
                'note_ref': correction_note_ref,
                'note_text': (
                    "The Entity corrected a prior period error and reflected the impact "
                    "in these financial statements."
                ),
            })

        if self.emphasis_liquidation:
            liquidation_note_ref = _next_emphasis_note_ref()
            emphasis_note_items.append({
                'key': 'liquidation',
                'label': 'Liquidation',
                'note_ref': liquidation_note_ref,
                'matter_text': "the Entity’s liquidation",
            })
            emphasis_note_lines.append({
                'note_ref': liquidation_note_ref,
                'note_text': (
                    "Management has resolved to liquidate the Entity and this matter has "
                    "been disclosed in these financial statements."
                ),
            })

        show_emphasis_of_matter = bool(emphasis_note_items)



        # Shareholders END #
        #########################################################################################################

        # Changes in Equity START #
        #########################################################################################################
        total_shares = statement_share_capital_total
        prev_total_shares = statement_prev_share_capital_total
        prev_prev_total_shares = statement_prev_prev_share_capital_total
        prior_opening_total_shares = statement_prior_opening_share_capital_total

        retained_earnings_code = '31010203'
        statutory_reserves_code = '31010301'
        owner_current_account_equity = _equity_head_balance(current_account_heads, owner_current_account_code)
        prev_owner_current_account_equity = _equity_head_balance(prev_account_heads, owner_current_account_code)
        prev_prev_owner_current_account_equity = _equity_head_balance(prev_prev_account_heads, owner_current_account_code)
        prior_opening_owner_current_account_equity = _equity_head_balance(
            prior_opening_account_heads,
            owner_current_account_code,
        )
        statutory_reserves_equity = _equity_head_balance(current_account_heads, statutory_reserves_code)
        prev_statutory_reserves_equity = _equity_head_balance(prev_account_heads, statutory_reserves_code)
        prev_prev_statutory_reserves_equity = _equity_head_balance(prev_prev_account_heads, statutory_reserves_code)
        prior_opening_statutory_reserves_equity = _equity_head_balance(
            prior_opening_account_heads,
            statutory_reserves_code,
        )
        prior_retained_earnings_balance = _equity_head_balance(prev_account_heads, retained_earnings_code)
        prev_prev_retained_earnings_balance = _equity_head_balance(prev_prev_account_heads, retained_earnings_code)
        prior_opening_retained_earnings_balance = _equity_head_balance(
            prior_opening_account_heads,
            retained_earnings_code,
        )
        retained_earnings = net_profit_after_tax
        prev_retained_earnings = prev_net_profit_after_tax
        prev_prev_retained_earnings = prev_prev_net_profit_after_tax
        if show_prior_year:
            # Keep prior-prior retained earnings from closing ledger balance.
            # Roll forward prior/current retained earnings with NPAT and dividends.
            prev_retained_earnings_balance = (
                prior_opening_retained_earnings_balance
                + prev_retained_earnings
                + prior_dividend_paid
            )
            retained_earnings_balance = (
                prev_retained_earnings_balance
                + retained_earnings
                + dividend_paid
            )
        else:
            prev_retained_earnings_balance = prior_retained_earnings_balance
            retained_earnings_balance = (
                prior_retained_earnings_balance
                + retained_earnings
                + dividend_paid
            )

        equity_total_display = (
            total_shares
            + owner_current_account_equity
            + retained_earnings_balance
            + statutory_reserves_equity
        )
        prev_equity_total_display = (
            prev_total_shares
            + prev_owner_current_account_equity
            + prev_retained_earnings_balance
            + prev_statutory_reserves_equity
        )
        prev_prev_equity_total_display = (
            prev_prev_total_shares
            + prev_prev_owner_current_account_equity
            + prev_prev_retained_earnings_balance
            + prev_prev_statutory_reserves_equity
        )
        prior_opening_equity_total_display = (
            prior_opening_total_shares
            + prior_opening_owner_current_account_equity
            + prior_opening_retained_earnings_balance
            + prior_opening_statutory_reserves_equity
        )
        balance_sheet_total_equity = equity_total_display
        prev_balance_sheet_total_equity = prev_equity_total_display

        # Total Equity to show on the statement of changes in equity
        total_soce = equity_total_display
        prev_total_soce = prev_equity_total_display
        prev_prev_total_soce = prev_prev_equity_total_display
        statutory_reserves_transfer = statutory_reserves_equity - prev_statutory_reserves_equity
        prior_statutory_reserves_transfer = (
            prev_statutory_reserves_equity - prior_opening_statutory_reserves_equity
        )
        total_of_equity_and_liabilities = balance_sheet_total_equity + total_liabilities
        prev_total_of_equity_and_liabilities = prev_balance_sheet_total_equity + prev_total_liabilities
        show_current_year_profit_line = True

        def _is_non_zero(value):
            # Keep row/column visibility aligned with displayed amounts,
            # which are rounded to whole numbers in the templates.
            return round(value or 0.0, 0) != 0

        def _format_date_label(date_value):
            return date_value.strftime('%d %B %Y') if date_value else ''

        def _total_comprehensive_label(amount, period_word):
            profit_or_loss = '(loss)' if (amount or 0.0) < 0 else 'profit'
            return f"Total comprehensive {profit_or_loss} for the {period_word}"

        def _period_word_for_range(start_date, end_date):
            if not start_date or not end_date:
                return 'period'
            expected_year_end = start_date + relativedelta(years=1) - relativedelta(days=1)
            return 'year' if expected_year_end == end_date else 'period'

        OWNED_PPE_SUBHEADS = [
            ('110101', 'Land and buildings'),
            ('110102', 'Furnitures and fixtures'),
            ('110103', 'Leasehold improvements'),
            ('110104', 'Vehicles'),
            ('110105', 'IT equipments'),
            ('110106', 'Office equipments'),
        ]
        owned_ppe_subhead_labels = {code: label for code, label in OWNED_PPE_SUBHEADS}
        owned_ppe_subhead_codes = [code for code, _label in OWNED_PPE_SUBHEADS]
        ppe_metric_keys = (
            'cost_opening',
            'cost_additions',
            'cost_disposals',
            'cost_closing',
            'acc_opening',
            'acc_charge',
            'acc_disposals',
            'acc_closing',
            'carrying_closing',
        )

        def _empty_ppe_metrics():
            return {
                'cost_opening': 0.0,
                'cost_additions': 0.0,
                'cost_disposals': 0.0,
                'cost_closing': 0.0,
                'acc_opening': 0.0,
                'acc_charge': 0.0,
                'acc_disposals': 0.0,
                'acc_closing': 0.0,
                'carrying_closing': 0.0,
            }

        def _owned_ppe_subhead_code(account_code):
            if not account_code or len(account_code) < 6:
                return None
            subhead_code = account_code[:6]
            if subhead_code not in owned_ppe_subhead_labels:
                return None
            return subhead_code

        def _is_accumulated_depreciation_row(row):
            name = (row.get('name') or '').lower()
            return (
                'accumulated depreciation' in name
                or 'acc depreciation' in name
            )

        def _build_ppe_schedule_data(start_date, end_date):
            metrics_by_subhead = {
                code: _empty_ppe_metrics()
                for code in owned_ppe_subhead_codes
            }
            opening_end_date = start_date - relativedelta(days=1) if start_date else False
            opening_rows = _fetch_account_rows(False, opening_end_date)
            period_rows = _fetch_account_rows(start_date, end_date)
            closing_rows = _fetch_account_rows(False, end_date)

            def _consume_rows(rows, value_loader):
                for row in rows:
                    subhead_code = _owned_ppe_subhead_code(row.get('code') or '')
                    if not subhead_code:
                        continue
                    subhead_metrics = metrics_by_subhead[subhead_code]
                    value_loader(subhead_metrics, row, _is_accumulated_depreciation_row(row))

            def _consume_opening(subhead_metrics, row, is_acc_dep):
                balance = row.get('balance', 0.0) or 0.0
                if is_acc_dep:
                    subhead_metrics['acc_opening'] += -balance
                else:
                    subhead_metrics['cost_opening'] += balance

            def _consume_period(subhead_metrics, row, is_acc_dep):
                debit = row.get('debit', 0.0) or 0.0
                credit = row.get('credit', 0.0) or 0.0
                if is_acc_dep:
                    subhead_metrics['acc_charge'] += credit
                    subhead_metrics['acc_disposals'] += -debit
                else:
                    subhead_metrics['cost_additions'] += debit
                    subhead_metrics['cost_disposals'] += -credit

            def _consume_closing(subhead_metrics, row, is_acc_dep):
                balance = row.get('balance', 0.0) or 0.0
                if is_acc_dep:
                    subhead_metrics['acc_closing'] += -balance
                else:
                    subhead_metrics['cost_closing'] += balance

            _consume_rows(opening_rows, _consume_opening)
            _consume_rows(period_rows, _consume_period)
            _consume_rows(closing_rows, _consume_closing)

            for code in owned_ppe_subhead_codes:
                subhead_metrics = metrics_by_subhead[code]
                subhead_metrics['carrying_closing'] = (
                    subhead_metrics['cost_closing'] - subhead_metrics['acc_closing']
                )

            return {
                'start_date': start_date,
                'end_date': end_date,
                'start_label': _format_date_label(start_date),
                'end_label': _format_date_label(end_date),
                'period_word': _period_word_for_range(start_date, end_date),
                'metrics': metrics_by_subhead,
            }

        ppe_schedule_data = []
        if date_start and date_end:
            ppe_schedule_data.append(_build_ppe_schedule_data(date_start, date_end))
        if show_prior_year and prior_date_start and prior_date_end:
            ppe_schedule_data.append(_build_ppe_schedule_data(prior_date_start, prior_date_end))

        ppe_visible_subheads = []
        for subhead_code in owned_ppe_subhead_codes:
            has_display_activity = any(
                _is_non_zero(
                    (schedule.get('metrics', {}).get(subhead_code, {}).get(metric_key, 0.0))
                )
                for schedule in ppe_schedule_data
                for metric_key in ppe_metric_keys
            )
            if has_display_activity:
                ppe_visible_subheads.append(subhead_code)

        if not ppe_visible_subheads:
            for subhead_code in owned_ppe_subhead_codes:
                has_raw_activity = any(
                    abs(schedule.get('metrics', {}).get(subhead_code, {}).get(metric_key, 0.0)) > 0.0
                    for schedule in ppe_schedule_data
                    for metric_key in ppe_metric_keys
                )
                if has_raw_activity:
                    ppe_visible_subheads.append(subhead_code)

        if not ppe_visible_subheads and ppe_schedule_data:
            ppe_visible_subheads = list(owned_ppe_subhead_codes)

        ppe_note_columns = [
            {'code': subhead_code, 'label': owned_ppe_subhead_labels[subhead_code]}
            for subhead_code in ppe_visible_subheads
        ]
        ppe_note_columns.append({'code': 'total', 'label': 'Total'})

        def _build_ppe_values(schedule_metrics, field_name):
            values = [schedule_metrics.get(code, {}).get(field_name, 0.0) for code in ppe_visible_subheads]
            total_value = sum(values)
            return values + [total_value]

        def _show_ppe_row(values):
            visible_values = values[:len(ppe_visible_subheads)]
            return any(_is_non_zero(value) for value in visible_values)

        ppe_note_schedules = []
        for schedule in ppe_schedule_data:
            schedule_metrics = schedule.get('metrics', {})
            schedule_rows = []
            section_values = [None] * len(ppe_note_columns)
            schedule_rows.append({
                'row_type': 'section',
                'label': 'Cost',
                'values': section_values,
            })
            schedule_rows.append({
                'row_type': 'line',
                'label': f"As at {schedule.get('start_label')}",
                'values': _build_ppe_values(schedule_metrics, 'cost_opening'),
            })
            additions_values = _build_ppe_values(schedule_metrics, 'cost_additions')
            if _show_ppe_row(additions_values):
                schedule_rows.append({
                    'row_type': 'line',
                    'label': f"Additions during the {schedule.get('period_word')}",
                    'values': additions_values,
                })
            disposals_values = _build_ppe_values(schedule_metrics, 'cost_disposals')
            if _show_ppe_row(disposals_values):
                schedule_rows.append({
                    'row_type': 'line',
                    'label': 'Disposals',
                    'values': disposals_values,
                })
            schedule_rows.append({
                'row_type': 'subtotal',
                'label': f"As at {schedule.get('end_label')}",
                'values': _build_ppe_values(schedule_metrics, 'cost_closing'),
            })
            schedule_rows.append({
                'row_type': 'section',
                'label': 'Accumulated depreciation',
                'values': [None] * len(ppe_note_columns),
            })
            schedule_rows.append({
                'row_type': 'line',
                'label': f"As at {schedule.get('start_label')}",
                'values': _build_ppe_values(schedule_metrics, 'acc_opening'),
            })
            charge_values = _build_ppe_values(schedule_metrics, 'acc_charge')
            if _show_ppe_row(charge_values):
                schedule_rows.append({
                    'row_type': 'line',
                    'label': f"Depreciation for the {schedule.get('period_word')}",
                    'values': charge_values,
                })
            acc_disposals_values = _build_ppe_values(schedule_metrics, 'acc_disposals')
            if _show_ppe_row(acc_disposals_values):
                schedule_rows.append({
                    'row_type': 'line',
                    'label': 'Disposals',
                    'values': acc_disposals_values,
                })
            schedule_rows.append({
                'row_type': 'subtotal',
                'label': f"As at {schedule.get('end_label')}",
                'values': _build_ppe_values(schedule_metrics, 'acc_closing'),
            })
            schedule_rows.append({
                'row_type': 'final',
                'label': f"Carrying value as at {schedule.get('end_label')}",
                'values': _build_ppe_values(schedule_metrics, 'carrying_closing'),
            })
            ppe_note_schedules.append({
                'start_label': schedule.get('start_label'),
                'end_label': schedule.get('end_label'),
                'period_word': schedule.get('period_word'),
                'rows': schedule_rows,
            })

        ppe_note_number = note_numbers.get('1101')

        prior_share_capital_movement = prev_total_shares - prior_opening_total_shares
        current_opening_share_capital = prev_total_shares if show_prior_year else prior_opening_total_shares
        current_share_capital_movement = total_shares - current_opening_share_capital

        soce_rows = []
        if show_prior_year and prior_date_start:
            soce_rows.append({
                'label': f"Balance as at {_format_date_label(soce_prior_opening_label_date)}",
                'share_capital': prior_opening_total_shares,
                'owner_current_account': prior_opening_owner_current_account_equity,
                'retained_earnings': prior_opening_retained_earnings_balance,
                'statutory_reserves': prior_opening_statutory_reserves_equity,
                'total_equity': prior_opening_equity_total_display,
                'is_balance': True,
            })
            if _is_non_zero(prior_share_capital_movement):
                soce_rows.append({
                    'label': "Share capital",
                    'share_capital': prior_share_capital_movement,
                    'owner_current_account': None,
                    'retained_earnings': None,
                    'statutory_reserves': None,
                    'total_equity': prior_share_capital_movement,
                    'is_balance': False,
                })
            if _is_non_zero(prior_owner_current_account):
                soce_rows.append({
                    'label': "Net movement in owner current\xa0account",
                    'share_capital': None,
                    'owner_current_account': prior_owner_current_account,
                    'retained_earnings': None,
                    'statutory_reserves': None,
                    'total_equity': prior_owner_current_account,
                    'is_balance': False,
                })
            if _is_non_zero(prior_statutory_reserves_transfer):
                soce_rows.append({
                    'label': "Transferred to statury reserves",
                    'share_capital': None,
                    'owner_current_account': None,
                    'retained_earnings': None,
                    'statutory_reserves': prior_statutory_reserves_transfer,
                    'total_equity': prior_statutory_reserves_transfer,
                    'is_balance': False,
                })
        if prior_date_end:
            if show_prior_year:
                prior_period_word = prior_column_period_word
                soce_rows.append({
                    'label': _total_comprehensive_label(prev_retained_earnings, prior_period_word),
                    'share_capital': None,
                    'owner_current_account': None,
                    'retained_earnings': prev_retained_earnings,
                    'statutory_reserves': None,
                    'total_equity': prev_retained_earnings,
                    'period_word': prior_period_word,
                    'is_balance': False,
                })
                if _is_non_zero(prior_dividend_paid):
                    soce_rows.append({
                        'label': "Dividend paid",
                        'share_capital': None,
                        'owner_current_account': None,
                        'retained_earnings': prior_dividend_paid,
                        'statutory_reserves': None,
                        'total_equity': prior_dividend_paid,
                        'is_balance': False,
                    })
            opening_current_date = date_start if not show_prior_year else prior_date_end
            if not opening_current_date:
                opening_current_date = prior_date_end
            opening_share_capital = None if not show_prior_year else prev_total_shares
            opening_owner_current_account = (
                None if not show_prior_year else prev_owner_current_account_equity
            )
            opening_retained_earnings = None if not show_prior_year else prev_retained_earnings_balance
            opening_statutory_reserves = None if not show_prior_year else prev_statutory_reserves_equity
            opening_total_equity = None if not show_prior_year else prev_equity_total_display
            soce_rows.append({
                'label': f"Balance as at {_format_date_label(opening_current_date)}",
                'share_capital': opening_share_capital,
                'owner_current_account': opening_owner_current_account,
                'retained_earnings': opening_retained_earnings,
                'statutory_reserves': opening_statutory_reserves,
                'total_equity': opening_total_equity,
                'is_balance': True,
            })
            if _is_non_zero(current_share_capital_movement):
                share_capital_value = current_share_capital_movement
                total_equity_value = current_share_capital_movement
                soce_rows.append({
                    'label': "Share capital",
                    'share_capital': share_capital_value,
                    'owner_current_account': None,
                    'retained_earnings': None,
                    'statutory_reserves': None,
                    'total_equity': total_equity_value,
                    'is_balance': False,
                })
            if _is_non_zero(owner_current_account):
                soce_rows.append({
                    'label': "Net movement in owner current\xa0account",
                    'share_capital': None,
                    'owner_current_account': owner_current_account,
                    'retained_earnings': None,
                    'statutory_reserves': None,
                    'total_equity': owner_current_account,
                    'is_balance': False,
                })
            if _is_non_zero(statutory_reserves_transfer):
                soce_rows.append({
                    'label': "Transferred to statury reserves",
                    'share_capital': None,
                    'owner_current_account': None,
                    'retained_earnings': None,
                    'statutory_reserves': statutory_reserves_transfer,
                    'total_equity': statutory_reserves_transfer,
                    'is_balance': False,
                })
        if show_current_year_profit_line:
            period_word = "period" if not show_prior_year else "year"
            current_profit_label = _total_comprehensive_label(retained_earnings, period_word)
            soce_rows.append({
                'label': current_profit_label,
                'share_capital': None,
                'owner_current_account': None,
                'retained_earnings': retained_earnings,
                'statutory_reserves': None,
                'total_equity': retained_earnings,
                'period_word': period_word,
                'is_balance': False,
            })
        if _is_non_zero(dividend_paid):
            soce_rows.append({
                'label': "Dividend paid",
                'share_capital': None,
                'owner_current_account': None,
                'retained_earnings': dividend_paid,
                'statutory_reserves': None,
                'total_equity': dividend_paid,
                'is_balance': False,
            })
        if date_end:
            soce_rows.append({
                'label': f"Balance as at {date_end.strftime('%d %B %Y')}",
                'share_capital': total_shares,
                'owner_current_account': owner_current_account_equity,
                'retained_earnings': retained_earnings_balance,
                'statutory_reserves': statutory_reserves_equity,
                'total_equity': equity_total_display,
                'is_balance': True,
            })

        # Show SOCE columns only when they contain non-zero displayed values.
        show_share_capital_column = any(
            row.get('share_capital') is not None and _is_non_zero(row.get('share_capital'))
            for row in soce_rows
        )
        show_owner_current_account_column = any(
            row.get('owner_current_account') is not None and _is_non_zero(row.get('owner_current_account'))
            for row in soce_rows
        )
        show_statutory_reserves_column = any(
            row.get('statutory_reserves') is not None and _is_non_zero(row.get('statutory_reserves'))
            for row in soce_rows
        )
        soce_column_count = (
            3
            + int(show_share_capital_column)
            + int(show_owner_current_account_column)
            + int(show_statutory_reserves_column)
        )
        show_owner_current_account_equity_row = (
            _is_non_zero(owner_current_account_equity)
            or (show_prior_year and _is_non_zero(prev_owner_current_account_equity))
        )
        show_owner_current_account_cashflow_row = (
            _is_non_zero(owner_current_account)
            or (show_prior_year and _is_non_zero(prior_owner_current_account))
        )

        # Changes in Equity END #
        #########################################################################################################

        # DMCC Supplementary Sheet START #
        #########################################################################################################
        dmcc_fixed_assets_net = sum(
            account.get('balance', 0.0)
            for account in current_account_heads.get('1101', {}).values()
            if 'accumulated depreciation' not in (account.get('name') or '').strip().lower()
        )
        dmcc_non_current_assets_excl_fixed_assets = non_current_assets_total - dmcc_fixed_assets_net
        dmcc_reserves = abs(current_prefix_totals[8].get('31010301', 0.0))
        dmcc_shareholders_current_account_loans = abs(
            current_prefix_totals[8].get(owner_current_account_code, 0.0)
        )
        dmcc_total_equity_system = balance_sheet_total_equity
        dmcc_total_salaries = (
            period_group_totals.get('5108', 0.0)
            + period_prefix_totals[8].get('51070101', 0.0)
        )
        dmcc_all_other_expenses = operating_expenses_total - dmcc_total_salaries
        dmcc_sheet_data = {
            'company_name': self.company_name or '',
            'portal_account_no': '',
            'customer_license_no': self.company_license_number or '',
            'year_start_date': date_start.strftime('%d/%m/%Y') if date_start else '',
            'year_end_date': date_end.strftime('%d/%m/%Y') if date_end else '',
            'total_share_capital': total_shares,
            'reserves': dmcc_reserves,
            'retained_earnings': retained_earnings_balance,
            'shareholders_current_account_loans': dmcc_shareholders_current_account_loans,
            'total_equity_system': dmcc_total_equity_system,
            'fixed_assets_net': dmcc_fixed_assets_net,
            'total_depreciation': depreciation_total,
            'current_assets': current_assets_total,
            'non_current_assets_excl_fixed_assets': dmcc_non_current_assets_excl_fixed_assets,
            'total_assets_system': total_assets,
            'current_liabilities': abs(current_liabilities_total),
            'non_current_liabilities': abs(non_current_liabilities_total),
            'total_liabilities_system': total_liabilities,
            'annual_sales_turnover': revenue_total,
            'cost_of_revenue_goods_sold': cost_of_revenue_total,
            'total_salaries': dmcc_total_salaries,
            'all_other_expenses': dmcc_all_other_expenses,
            'all_other_income': other_income_total,
            'gross_profit_loss_system': gross_profit,
            'net_profit_loss_system': net_profit_after_tax,
            'audit_firm_name': 'Assurance Corp Audit and Accounting',
            'auditor_signature': '',
            'auditor_date': generated_date_display,
            'auditor_seal': '',
        }
        # DMCC Supplementary Sheet END #
        #########################################################################################################

        # Statement of Cash Flows START #
        #########################################################################################################


        


        # Statement of Cash Flows END #
        #########################################################################################################

        tb_diff_current = total_assets - total_of_equity_and_liabilities
        tb_diff_prior = prev_total_assets - prev_total_of_equity_and_liabilities if show_prior_year else 0.0
        tb_warning_current = self._compose_tb_warning('Current period', tb_diff_current)
        tb_warning_prior = (
            self._compose_tb_warning('Prior period', tb_diff_prior)
            if show_prior_year else False
        )
        self.tb_diff_current = tb_diff_current
        self.tb_diff_prior = tb_diff_prior
        self.tb_warning_current = tb_warning_current or False
        self.tb_warning_prior = tb_warning_prior or False
        self._sync_tb_overrides_json()
        if tb_warning_current or tb_warning_prior:
            _logger.warning(
                (
                    "AUDIT_TB_MISMATCH wizard_id=%s company_id=%s "
                    "current_diff=%.6f prior_diff=%.6f"
                ),
                self.id,
                self.company_id.id,
                tb_diff_current,
                tb_diff_prior,
            )

        return {

            # Balance Sheet
            'balance_rows': balance_rows,
            'profit_loss_accounts': profit_loss_accounts,
            'total_assets': total_assets,
            'total_liabilities': total_liabilities,
            'total_equity': balance_sheet_total_equity,
            'current_assets': current_assets,
            'current_liabilities': current_liabilities,
            'equity': equity,
            'non_current_assets': non_current_assets,
            'non_current_liabilities': non_current_liabilities,
            'current_liabilities_total': current_liabilities_total,
            'non_current_liabilities_total': non_current_liabilities_total,
            'current_assets_total': current_assets_total,
            'non_current_assets_total': non_current_assets_total,
            'total_of_equity_and_liabilities': total_of_equity_and_liabilities,
            'tb_diff_current': tb_diff_current,
            'tb_warning_current': tb_warning_current,
            'prev_current_assets_total': prev_current_assets_total,
            'prev_non_current_assets_total': prev_non_current_assets_total,
            'prev_total_assets': prev_total_assets,
            'prev_current_liabilities_total': prev_current_liabilities_total,
            'prev_non_current_liabilities_total': prev_non_current_liabilities_total,
            'prev_total_liabilities': prev_total_liabilities,
            'prev_total_equity': prev_balance_sheet_total_equity,
            'prev_total_of_equity_and_liabilities': prev_total_of_equity_and_liabilities,
            'tb_diff_prior': tb_diff_prior,
            'tb_warning_prior': tb_warning_prior,

            # PnL
             'cost_of_revenue_total':cost_of_revenue_total,
             'depreciation_total':depreciation_total,
             'operating_expenses_total':operating_expenses_total,
             'revenue_total':revenue_total,
             'gross_profit':gross_profit,
             'net_profit_before_tax':net_profit_before_tax,
             'net_profit_after_tax':net_profit_after_tax,
             'other_income_total': other_income_total,
             'gross_profit_margin': gross_profit_margin,
             'gross_profit_margin_label': gross_profit_margin_label,
             'net_profit_margin': net_profit_margin,
             'tax_amount':tax_amount,
             'prev_revenue_total': prev_revenue_total,
             'prev_cost_of_revenue_total': prev_cost_of_revenue_total,
             'prev_gross_profit': prev_gross_profit,
             'prev_operating_expenses_total': prev_operating_expenses_total,
             'prev_net_profit_before_tax': prev_net_profit_before_tax,
             'prev_tax_amount': prev_tax_amount,
             'prev_net_profit_after_tax': prev_net_profit_after_tax,
             'prev_other_income_total': prev_other_income_total,
             'prev_gross_profit_margin': prev_gross_profit_margin,
             'prev_net_profit_margin': prev_net_profit_margin,
             'net_profit_before_tax_label': net_profit_before_tax_label,
             'net_profit_after_tax_label': net_profit_after_tax_label,

             # Statement of Changes in Equity
             'retained_earnings':retained_earnings,
            'total_soce':total_soce,
            'prev_retained_earnings': prev_retained_earnings,
            'prev_total_soce': prev_total_soce,
            'prev_prev_retained_earnings': prev_prev_retained_earnings,
            'prev_prev_total_soce': prev_prev_total_soce,
            'share_capital_total': total_shares,
            'prev_share_capital_total': prev_total_shares,
            'owner_current_account_equity': owner_current_account_equity,
            'prev_owner_current_account_equity': prev_owner_current_account_equity,
            'statutory_reserves_equity': statutory_reserves_equity,
            'prev_statutory_reserves_equity': prev_statutory_reserves_equity,
            'retained_earnings_balance': retained_earnings_balance,
            'prev_retained_earnings_balance': prev_retained_earnings_balance,
            'prev_prev_retained_earnings_balance': prev_prev_retained_earnings_balance,
            'equity_total_display': equity_total_display,
            'prev_equity_total_display': prev_equity_total_display,
            'prev_prev_equity_total_display': prev_prev_equity_total_display,
            'soce_rows': soce_rows,
            'show_share_capital_column': show_share_capital_column,
            'show_owner_current_account_column': show_owner_current_account_column,
            'show_statutory_reserves_column': show_statutory_reserves_column,
            'soce_column_count': soce_column_count,
            'show_owner_current_account_equity_row': show_owner_current_account_equity_row,
            'show_owner_current_account_cashflow_row': show_owner_current_account_cashflow_row,
            'comparative_period_word': comparative_period_word,
            'signature_names': signature_names,
            'signature_display_names': signature_display_names,
            'signature_entries': signature_entries,
            'owner': '',
            'owner_display_names': owner_display_names,
            'shareholder_display_text': shareholder_display_text,
            'management_director_text': management_director_text,
            'management_director_title': management_director_title,
            'generated_date_display': generated_date_display,
            'share_rows': share_rows,
            'authorized_share_capital': authorized_share_capital,
            'total_shares_count': total_shares_count,
            'share_value_default': share_value_default,
            'share_capital_paid_status': self.share_capital_paid_status,
            'show_share_capital_conversion_note': self.show_share_capital_conversion_note,
            'share_conversion_currency': (self.share_conversion_currency or '').strip(),
            'share_conversion_original_value': self.share_conversion_original_value or 0.0,
            'share_conversion_exchange_rate': self.share_conversion_exchange_rate or 0.0,
            'show_share_capital_transfer_note': self.show_share_capital_transfer_note,
            'share_transfer_date': self.share_transfer_date,
            'share_transfer_from': (self.share_transfer_from or '').strip(),
            'share_transfer_shares': self.share_transfer_shares or 0,
            'share_transfer_percentage': self.share_transfer_percentage or 0.0,
            'share_transfer_to': (self.share_transfer_to or '').strip(),
            'related_party_name': self.related_party_name,
            'related_party_relationship': self.related_party_relationship,
            'related_party_transaction': self.related_party_transaction,
            'related_party_amount': self.related_party_amount,
            'related_party_amount_prior': self.related_party_amount_prior,
            'show_related_parties_note': self.show_related_parties_note,
            'show_shareholder_note': self.show_shareholder_note,
            'corporate_tax_liability_paid': self.corporate_tax_liability_paid,
            'show_ct_first_tax_year_line': self.show_ct_first_tax_year_line,
            'gap_report_of_directors': self.gap_report_of_directors,
            'gap_independent_auditor_report': self.gap_independent_auditor_report,
            'gap_notes_to_financial_statements': self.gap_notes_to_financial_statements,
            'signature_break_lines_report_of_directors': max(
                int(self.signature_break_lines_report_of_directors or 0), 0
            ),
            'signature_break_lines_balance_sheet': max(int(self.signature_break_lines_balance_sheet or 0), 0),
            'signature_break_lines_profit_loss': max(int(self.signature_break_lines_profit_loss or 0), 0),
            'signature_break_lines_changes_in_equity': max(
                int(self.signature_break_lines_changes_in_equity or 0), 0
            ),
            'signature_break_lines_cash_flows': max(int(self.signature_break_lines_cash_flows or 0), 0),
            'signature_break_lines_notes': max(int(self.signature_break_lines_notes or 0), 0),

            # Shareholders
            'sh_name_1':sh_name_1,
             'sh_nat_1':sh_nat_1,
             'sh_shares_1':sh_shares_1,
             'sh_value_1':sh_value_1,
             'sh_total_1':sh_total_1,
             'sh_name_2':sh_name_2,
             'sh_nat_2':sh_nat_2,
             'sh_shares_2':sh_shares_2,
             'sh_value_2':sh_value_2,
             'sh_total_2':sh_total_2,
             'sh_name_3':sh_name_3,
             'sh_nat_3':sh_nat_3,
             'sh_shares_3':sh_shares_3,
             'sh_value_3':sh_value_3,
             'sh_total_3':sh_total_3,

             # Property and equipment subheads
             'current_ppe_subheads': current_ppe_subheads,
             'current_ppe_total': current_ppe_total,
             'prev_ppe_subheads': prev_ppe_subheads,
             'prev_ppe_total': prev_ppe_total,
             'current_subheads_by_group': current_subheads_by_group,
             'prev_subheads_by_group': prev_subheads_by_group,
             'current_main_heads': current_main_heads,
             'prev_main_heads': prev_main_heads,
            'main_head_labels': MAIN_HEAD_LABELS,
            'current_group_totals': current_group_totals,
            'prev_group_totals': prev_group_totals,
            'show_prior_year': show_prior_year,
            'audit_period_category': period_category,
            'balance_sheet_date_mode': self.balance_sheet_date_mode,
            'prior_balance_sheet_date_mode': self.prior_balance_sheet_date_mode,
            'is_dormant_period': period_category.startswith('dormant_'),
            'is_cessation_period': period_category.startswith('cessation_'),
            'fa_other_receivables_current': other_receivables_current,
            'fa_other_receivables_prev': other_receivables_prev,

            # Cash flow
            'current_depreciation_total': current_depreciation_total,
            'prior_depreciation_total': prior_depreciation_total,
            'end_service_benefits_adjustment': end_service_benefits_adjustment,
            'prior_end_service_benefits_adjustment': prior_end_service_benefits_adjustment,
            'operating_cashflow_before_working_capital': operating_cashflow_before_working_capital,
            'prior_operating_cashflow_before_working_capital': prior_operating_cashflow_before_working_capital,
            'change_in_current_assets': change_in_current_assets,
            'prior_change_in_current_assets': prior_change_in_current_assets,
            'change_in_current_liabilities': change_in_current_liabilities,
            'prior_change_in_current_liabilities': prior_change_in_current_liabilities,
            'corporate_tax_paid': corporate_tax_paid,
            'prior_corporate_tax_paid': prior_corporate_tax_paid,
            'end_service_benefits_paid': end_service_benefits_paid,
            'prior_end_service_benefits_paid': prior_end_service_benefits_paid,
            'net_cash_generated_from_operations': net_cash_generated_from_operations,
            'prior_net_cash_generated_from_operations': prior_net_cash_generated_from_operations,
            'current_property': current_property,
            'prior_property': prior_property,
            'net_cash_generated_from_investing_activities': net_cash_generated_from_investing_activities,
            'prior_net_cash_generated_from_investing_activities': prior_net_cash_generated_from_investing_activities,
            'paid_up_capital': paid_up_capital,
            'prior_paid_up_capital': prior_paid_up_capital,
            'dividend_paid': dividend_paid,
            'prior_dividend_paid': prior_dividend_paid,
            'owner_current_account': owner_current_account,
            'prior_owner_current_account': prior_owner_current_account,
            'net_cash_generated_from_financing_activities': net_cash_generated_from_financing_activities,
            'prior_net_cash_generated_from_financing_activities': prior_net_cash_generated_from_financing_activities,
            'net_cash_and_cash_equivalents': net_cash_and_cash_equivalents,
            'prior_net_cash_and_cash_equivalents': prior_net_cash_and_cash_equivalents,
            'cash_beginning_year': cash_beginning_year,
            'cash_end_of_year': cash_end_of_year,
            'prior_cash_beginning_year': prior_cash_beginning_year,
            'prior_cash_end_of_year': prior_cash_end_of_year,

            # Notes
            'show_emphasis_of_matter': show_emphasis_of_matter,
            'emphasis_note_items': emphasis_note_items,
            'emphasis_note_lines': emphasis_note_lines,
            'ppe_note_number': ppe_note_number,
            'ppe_note_columns': ppe_note_columns,
            'ppe_note_schedules': ppe_note_schedules,
            'note_numbers': note_numbers,
            'note_sections': note_sections,
            'last_note_number': last_note_number,
            'dmcc_sheet_data': dmcc_sheet_data,
            }


    @staticmethod
    def _normalize_account_code(code):
        return ''.join(ch for ch in (code or '') if ch.isdigit())

    @staticmethod
    def _to_float(value):
        try:
            return float(value or 0.0)
        except (TypeError, ValueError):
            return 0.0

    @staticmethod
    def _sum_prefix_balances(balance_by_code, *prefixes):
        return sum(
            amount
            for code, amount in (balance_by_code or {}).items()
            if any((code or '').startswith(prefix) for prefix in prefixes)
        )

    @staticmethod
    def _clear_rows(ws, start_row, end_row, min_col, max_col):
        for row in range(start_row, end_row + 1):
            for col in range(min_col, max_col + 1):
                ws.cell(row=row, column=col, value=None)

    def action_print_account_report_lines(self):
        """Generate and display report"""
        self.ensure_one()
        self._validate_emphasis_options()

        if self.use_previous_settings:
            self._store_previous_settings()

        # report_data = self._get_report_data()
        # You can now pass `report_data` to your Word/PDF renderer

        return {
            'type': 'ir.actions.act_url',
            'url': f'/audit_report/view/{self.id}',
            'target': 'new',
        } 

    def _get_saved_document_name(self):
        self.ensure_one()
        report_type_label = dict(self._fields['report_type'].selection).get(self.report_type, self.report_type or '')
        company_name = self.company_id.name or 'Company'
        end_date_label = self.date_end.strftime('%d %B %Y') if self.date_end else ''
        name = f"{company_name} - {report_type_label} {end_date_label}".strip()
        return name or f"Audit Report {self.id}"

    def _get_wizard_snapshot_json(self):
        self.ensure_one()
        tb_overrides_json = self._sync_tb_overrides_json()
        snapshot = {
            'wizard_id': self.id,
            'company_id': self.company_id.id,
            'date_start': self.date_start.isoformat() if self.date_start else False,
            'date_end': self.date_end.isoformat() if self.date_end else False,
            'balance_sheet_date_mode': self.balance_sheet_date_mode,
            'prior_balance_sheet_date_mode': self.prior_balance_sheet_date_mode,
            'report_type': self.report_type,
            'audit_period_category': self.audit_period_category,
            'auditor_type': self.auditor_type,
            'prior_year_mode': self.prior_year_mode,
            'prior_date_start': self.prior_date_start.isoformat() if self.prior_date_start else False,
            'prior_date_end': self.prior_date_end.isoformat() if self.prior_date_end else False,
            'soce_prior_opening_label_date': (
                self.soce_prior_opening_label_date.isoformat()
                if self.soce_prior_opening_label_date else False
            ),
            'show_related_parties_note': self.show_related_parties_note,
            'show_shareholder_note': self.show_shareholder_note,
            'corporate_tax_liability_paid': self.corporate_tax_liability_paid,
            'show_ct_first_tax_year_line': self.show_ct_first_tax_year_line,
            'share_capital_paid_status': self.share_capital_paid_status,
            'show_share_capital_conversion_note': self.show_share_capital_conversion_note,
            'share_conversion_currency': self.share_conversion_currency,
            'share_conversion_original_value': self.share_conversion_original_value,
            'share_conversion_exchange_rate': self.share_conversion_exchange_rate,
            'show_share_capital_transfer_note': self.show_share_capital_transfer_note,
            'share_transfer_date': self.share_transfer_date.isoformat() if self.share_transfer_date else False,
            'share_transfer_from': self.share_transfer_from,
            'share_transfer_shares': self.share_transfer_shares,
            'share_transfer_percentage': self.share_transfer_percentage,
            'share_transfer_to': self.share_transfer_to,
            'tb_include_zero_accounts': bool(self.tb_include_zero_accounts),
            'tb_overrides_json': tb_overrides_json,
            'tb_diff_current': self._to_float(self.tb_diff_current),
            'tb_diff_prior': self._to_float(self.tb_diff_prior),
            'tb_warning_current': self.tb_warning_current or False,
            'tb_warning_prior': self.tb_warning_prior or False,
            'generated_by_user_id': self.env.user.id,
        }
        for i in range(1, 11):
            snapshot[f'signature_include_{i}'] = bool(getattr(self, f'signature_include_{i}'))
            snapshot[f'signature_role_{i}'] = getattr(self, f'signature_role_{i}') or (
                'primary' if i == 1 else 'secondary'
            )
            snapshot[f'director_include_{i}'] = bool(getattr(self, f'director_include_{i}'))
            snapshot[f'owner_include_{i}'] = bool(getattr(self, f'owner_include_{i}'))
        return json.dumps(snapshot)

    def action_save_editable_report(self):
        """Create a company-scoped saved report document and initial revision."""
        self.ensure_one()
        self._validate_emphasis_options()
        flow_started = time.perf_counter()

        if self.use_previous_settings:
            self._store_previous_settings()

        from ..controllers.main import AuditReportController

        controller = AuditReportController()
        templates_path = controller._templates_path()
        template_env = controller._get_template_env(templates_path)
        css_content = controller._get_cached_css_content(controller._css_path())

        report_data_started = time.perf_counter()
        report_data = self._get_report_data()
        _logger.debug(
            "AUDIT_PERF action_save_editable_report wizard_id=%s stage=report_data elapsed_ms=%.2f",
            self.id,
            (time.perf_counter() - report_data_started) * 1000.0,
        )

        toc_started = time.perf_counter()
        try:
            toc_entries = controller._compute_toc_entries(
                self,
                report_data=report_data,
                template_env=template_env,
                css_content=css_content,
            )
        except Exception:
            toc_entries = None
        _logger.debug(
            "AUDIT_PERF action_save_editable_report wizard_id=%s stage=toc elapsed_ms=%.2f",
            self.id,
            (time.perf_counter() - toc_started) * 1000.0,
        )

        html_started = time.perf_counter()
        html_with_style = controller._render_report_html(
            self,
            toc_entries=toc_entries,
            report_data=report_data,
            template_env=template_env,
            css_content=css_content,
        )
        _logger.debug(
            "AUDIT_PERF action_save_editable_report wizard_id=%s stage=html elapsed_ms=%.2f",
            self.id,
            (time.perf_counter() - html_started) * 1000.0,
        )
        if not html_with_style:
            raise ValidationError("Unable to render report HTML for saving.")
        tb_overrides_json = self._sync_tb_overrides_json()

        document = self.env['audit.report.document'].create({
            'name': self._get_saved_document_name(),
            'company_id': self.company_id.id,
            'date_start': self.date_start,
            'date_end': self.date_end,
            'report_type': self.report_type,
            'audit_period_category': self.audit_period_category,
            'source_wizard_json': self._get_wizard_snapshot_json(),
            'tb_overrides_json': tb_overrides_json,
        })
        revision = document.create_revision_from_html(
            html_with_style,
            tb_overrides_json=tb_overrides_json,
        )
        pdf_started = time.perf_counter()
        generate_initial_pdf = bool(self.env.context.get('audit_generate_initial_pdf'))
        if generate_initial_pdf:
            revision._get_pdf_content(pre_rendered_html=html_with_style, refresh_toc=False)
        pdf_elapsed_ms = (time.perf_counter() - pdf_started) * 1000.0
        _logger.debug(
            "AUDIT_PERF action_save_editable_report wizard_id=%s stage=pdf elapsed_ms=%.2f skipped=%s total_elapsed_ms=%.2f",
            self.id,
            pdf_elapsed_ms,
            not generate_initial_pdf,
            (time.perf_counter() - flow_started) * 1000.0,
        )

        return {
            'type': 'ir.actions.act_window',
            'name': 'Saved Audit Report',
            'res_model': 'audit.report.document',
            'view_mode': 'form',
            'res_id': document.id,
            'target': 'current',
        }

    def action_print_account_report_pdf(self):
        """Generate and display report as server-side PDF (headless Chrome)."""
        self.ensure_one()
        self._validate_emphasis_options()

        if self.use_previous_settings:
            self._store_previous_settings()

        return {
            'type': 'ir.actions.act_url',
            'url': f'/audit_report/pdf/{self.id}',
            'target': 'new',
        }

    
    
