# -*- coding: utf-8 -*-

from collections import defaultdict

from odoo import models, fields, api, _, SUPERUSER_ID
from odoo.exceptions import AccessError
import logging

_logger = logging.getLogger(__name__)


class VatThreshold(models.Model):
    _name = 'vat.threshold'
    _description = 'VAT Threshold Management'
    _order = 'last_check_date desc'
    _sql_constraints = [
        (
            'vat_threshold_company_unique',
            'unique(company_id)',
            'A VAT threshold record already exists for this company.',
        ),
    ]

    currency_id = fields.Many2one(
        'res.currency',
        related='company_id.currency_id',
        readonly=True,
        store=True
    )
    company_id = fields.Many2one(
        'res.company',
        string='Company',
        required=True,
        ondelete='cascade',
        # index=True,
        # check_company=True,
    )
    company_name = fields.Char(
        string='Company Name',
        related='company_id.name',
        store=True,
        readonly=True
    )
    
    # Revenue Account Codes
    revenue_uae_clients = fields.Monetary(
        string='Revenue - UAE Clients (41010101)',
        compute='_compute_revenue',
        store=True,
        currency_field='currency_id',
        help='Account code 41010101'
    )
    revenue_international_clients = fields.Monetary(
        string='Revenue - International Clients (41010102)',
        compute='_compute_revenue',
        store=True,
        currency_field='currency_id',
        help='Account code 41010102'
    )
    revenue_related_party = fields.Monetary(
        string='Revenue - Related Party (41020101)',
        compute='_compute_revenue',
        store=True,
        currency_field='currency_id',
        help='Account code 41020101'
    )
    
    total_sales = fields.Monetary(
        string='Total Sales',
        compute='_compute_revenue',
        store=True,
        currency_field='currency_id',
        help='Sum of all revenue accounts'
    )
    
    is_vat_registered = fields.Boolean(
        string='VAT Registered',
        compute='_compute_vat_status',
        store=True,
        help='Checked if company has VAT registration number'
    )
    
    email_sent = fields.Boolean(
        string='Email Sent',
        default=False,
        help='Indicates if threshold notification email was sent'
    )
    
    vat_verified = fields.Boolean(
        string='VAT Verified',
        default=False,
        help='Manually mark if VAT status is verified'
    )
    
    threshold_action = fields.Selection(
        [
            ('none', 'No Action Required'),
            ('register', 'Register for VAT'),
            ('cancel', 'Cancel VAT Registration'),
        ],
        string='Action Required',
        default='none',
        compute='_compute_threshold_action',
        store=True
    )
    
    last_check_date = fields.Datetime(
        string='Last Check Date',
        default=fields.Datetime.now
    )
    
    email_sent_date = fields.Datetime(
        string='Email Sent Date',
        readonly=True
    )
    
    # Rolling Period Fields
    date_from = fields.Date(
        string='Date From',
        compute='_compute_rolling_period',
        store=True,
        help='Start date of 9-month rolling period'
    )
    date_to = fields.Date(
        string='Date To',
        compute='_compute_rolling_period',
        store=True,
        help='End date of rolling period (today)'
    )
    rolling_period_months = fields.Integer(
        string='Rolling Period (Months)',
        default=9,
        help='Number of months for rolling period calculation'
    )
    
    # Rolling Period Revenue (9 months)
    rolling_revenue_uae_clients = fields.Monetary(
        string='Rolling Revenue - UAE Clients',
        compute='_compute_rolling_revenue',
        store=True,
        currency_field='currency_id',
        help='UAE Clients revenue for rolling 9-month period'
    )
    rolling_revenue_international_clients = fields.Monetary(
        string='Rolling Revenue - International Clients',
        compute='_compute_rolling_revenue',
        store=True,
        currency_field='currency_id',
        help='International Clients revenue for rolling 9-month period'
    )
    rolling_revenue_related_party = fields.Monetary(
        string='Rolling Revenue - Related Party',
        compute='_compute_rolling_revenue',
        store=True,
        currency_field='currency_id',
        help='Related Party revenue for rolling 9-month period'
    )
    rolling_total_sales = fields.Monetary(
        string='Rolling Total Sales (9 Months)',
        compute='_compute_rolling_revenue',
        store=True,
        currency_field='currency_id',
        help='Total sales for rolling 9-month period'
    )

    @api.depends('company_id')
    def _compute_revenue(self):
        for record in self:
            if not record.company_id:
                record.revenue_uae_clients = 0.0
                record.revenue_international_clients = 0.0
                record.revenue_related_party = 0.0
                record.total_sales = 0.0
                continue
            
            # Get account move lines for revenue accounts
            # Account codes: 41010101, 41010102, 41020101
            account_codes = ['41010101', '41010102', '41020101']
            
            # Search in the target company context so company-dependent codes resolve correctly.
            account_env = self.env['account.account'].with_company(record.company_id).sudo()
            accounts = account_env.search([
                ('code', 'in', account_codes),
                ('company_ids', 'in', [record.company_id.id])
            ])
            
            # Initialize values
            record.revenue_uae_clients = 0.0
            record.revenue_international_clients = 0.0
            record.revenue_related_party = 0.0
            
            balance_by_account = self._get_account_balances(record.company_id, accounts.ids)

            balance_by_code = defaultdict(float)
            for account in accounts:
                balance_by_code[account.code] += balance_by_account.get(account.id, 0.0)

            record.revenue_uae_clients = balance_by_code['41010101']
            record.revenue_international_clients = balance_by_code['41010102']
            record.revenue_related_party = balance_by_code['41020101']
            
            record.total_sales = (
                record.revenue_uae_clients + 
                record.revenue_international_clients + 
                record.revenue_related_party
            )

    @api.depends('company_id', 'company_id.vat_registration_number')
    def _compute_vat_status(self):
        for record in self:
            # Check if company has VAT registration number (from company_exte module)
            record.is_vat_registered = bool(record.company_id.vat_registration_number)

    @api.depends('last_check_date', 'rolling_period_months')
    def _compute_rolling_period(self):
        """Calculate rolling period dates (default 9 months)"""
        from dateutil.relativedelta import relativedelta
        
        for record in self:
            # Use today as end date
            date_to = fields.Date.today()
            # Calculate start date by subtracting months
            date_from = date_to - relativedelta(months=record.rolling_period_months)
            
            record.date_to = date_to
            record.date_from = date_from

    @api.depends('company_id', 'date_from', 'date_to', 'rolling_period_months')
    def _compute_rolling_revenue(self):
        """Calculate revenue for rolling 9-month period"""
        from dateutil.relativedelta import relativedelta
        
        for record in self:
            if not record.company_id:
                record.rolling_revenue_uae_clients = 0.0
                record.rolling_revenue_international_clients = 0.0
                record.rolling_revenue_related_party = 0.0
                record.rolling_total_sales = 0.0
                continue
            
            # Ensure dates are computed
            if not record.date_from or not record.date_to:
                date_to = fields.Date.today()
                date_from = date_to - relativedelta(months=record.rolling_period_months or 9)
            else:
                date_from = record.date_from
                date_to = record.date_to
            
            _logger.info("Computing rolling revenue for company: %s, date_from: %s, date_to: %s", 
                        record.company_id.name, date_from, date_to)
            
            # Account codes for revenue - using LIKE for flexible matching
            # 41010101 = UAE Clients, 41010102 = International Clients, 41020101 = Related Party
            account_env = self.env['account.account'].with_company(record.company_id).sudo()
            domain_uae = [
                ('code', '=like', '41010101%'),
                ('company_ids', 'in', [record.company_id.id])
            ]
            domain_intl = [
                ('code', '=like', '41010102%'),
                ('company_ids', 'in', [record.company_id.id])
            ]
            domain_related = [
                ('code', '=like', '41020101%'),
                ('company_ids', 'in', [record.company_id.id])
            ]
            
            # Initialize values
            record.rolling_revenue_uae_clients = 0.0
            record.rolling_revenue_international_clients = 0.0
            record.rolling_revenue_related_party = 0.0
            
            accounts_uae = account_env.search(domain_uae)
            record.rolling_revenue_uae_clients = self._get_account_balance(
                record.company_id,
                accounts_uae.ids,
                date_from=date_from,
                date_to=date_to,
            )
            _logger.info("UAE rolling revenue for %s: %s", record.company_id.name, record.rolling_revenue_uae_clients)

            accounts_intl = account_env.search(domain_intl)
            record.rolling_revenue_international_clients = self._get_account_balance(
                record.company_id,
                accounts_intl.ids,
                date_from=date_from,
                date_to=date_to,
            )
            _logger.info("Intl rolling revenue for %s: %s", record.company_id.name, record.rolling_revenue_international_clients)

            accounts_related = account_env.search(domain_related)
            record.rolling_revenue_related_party = self._get_account_balance(
                record.company_id,
                accounts_related.ids,
                date_from=date_from,
                date_to=date_to,
            )
            _logger.info("Related rolling revenue for %s: %s", record.company_id.name, record.rolling_revenue_related_party)
            
            record.rolling_total_sales = (
                record.rolling_revenue_uae_clients + 
                record.rolling_revenue_international_clients + 
                record.rolling_revenue_related_party
            )
            _logger.info("Total rolling sales for %s: %s", record.company_id.name, record.rolling_total_sales)

    @api.depends('rolling_total_sales', 'is_vat_registered')
    def _compute_threshold_action(self):
        """
        Logic based on 9-month rolling period sales:
        - VAT Registered + Sales < 125,000 → Cancel VAT
        - Not VAT Registered + Sales > 375,000 → Register VAT
        """
        MIN_THRESHOLD = 125000  # For VAT cancellation
        MAX_THRESHOLD = 375000  # For VAT registration
        
        for record in self:
            # Use rolling_total_sales (9-month period) for threshold calculation
            sales = record.rolling_total_sales or 0.0
            
            if record.is_vat_registered:
                # If VAT registered and sales below minimum threshold
                if sales < MIN_THRESHOLD:
                    record.threshold_action = 'cancel'
                else:
                    record.threshold_action = 'none'
            else:
                # If not VAT registered and sales above maximum threshold
                if sales > MAX_THRESHOLD:
                    record.threshold_action = 'register'
                else:
                    record.threshold_action = 'none'

    def _get_account_balances(self, company, account_ids, date_from=None, date_to=None):
        """Return a mapping of account id -> balance for posted journal items."""
        if not account_ids:
            return {}

        domain = [
            ('account_id', 'in', account_ids),
            ('move_id.state', '=', 'posted'),
            ('company_id', '=', company.id),
        ]
        if date_from:
            domain.append(('move_id.date', '>=', date_from))
        if date_to:
            domain.append(('move_id.date', '<=', date_to))

        grouped = self.env['account.move.line'].with_company(company).sudo().read_group(
            domain,
            ['account_id', 'credit:sum', 'debit:sum'],
            ['account_id'],
            lazy=False,
        )
        return {
            row['account_id'][0]: (row.get('credit', 0.0) or 0.0) - (row.get('debit', 0.0) or 0.0)
            for row in grouped
            if row.get('account_id')
        }

    def _get_account_balance(self, company, account_ids, date_from=None, date_to=None):
        """Return the total balance for a set of accounts (sum of all account balances)."""
        balances = self._get_account_balances(company, account_ids, date_from=date_from, date_to=date_to)
        return sum(balances.values()) if balances else 0.0

    def _ensure_system_user(self):
        if self.env.uid != SUPERUSER_ID and not self.env.user.has_group('base.group_system'):
            raise AccessError(_('This action is restricted to system administrators.'))

    def _refresh_threshold_values(self):
        """Recompute all stored VAT threshold values for the current recordset."""
        for record in self:
            previous_action = record.threshold_action
            record._compute_rolling_period()
            record._compute_rolling_revenue()
            record._compute_revenue()
            record._compute_vat_status()
            record._compute_threshold_action()
            vals = {'last_check_date': fields.Datetime.now()}
            if record.threshold_action != previous_action:
                vals.update({
                    'email_sent': False,
                    'email_sent_date': False,
                })
            record.write(vals)

    def action_check_threshold(self):
        """Manual action to check threshold and send email if needed"""
        self._ensure_system_user()
        self._refresh_threshold_values()
        
        # Send email if action required and not already sent
        for record in self:
            if record.threshold_action != 'none' and not record.email_sent:
                record.send_threshold_notification()
        
        return {
            'type': 'ir.actions.client',
            'tag': 'reload',
        }

    def send_threshold_notification(self):
        """Send email notification based on threshold action"""
        self.ensure_one()
        
        if self.threshold_action == 'none':
            return False
        
        # Find the appropriate email template from the module XML IDs.
        template_map = {
            'register': 'vat_threshold.email_template_vat_register',
            'cancel': 'vat_threshold.email_template_vat_cancel',
        }
        template_xmlid = template_map.get(self.threshold_action)
        template = self.env.ref(template_xmlid, raise_if_not_found=False) if template_xmlid else False
        
        if not template:
            _logger.warning("Email template not found for action: %s", self.threshold_action)
            return False
        
        # Send the email
        try:
            template.send_mail(self.id, force_send=True)
            self.write({
                'email_sent': True,
                'email_sent_date': fields.Datetime.now()
            })
            _logger.info("VAT threshold notification sent for company: %s", self.company_id.name)
            return True
        except Exception as e:
            _logger.error("Failed to send VAT threshold email: %s", str(e))
            return False

    @api.model
    def _cron_check_vat_threshold(self):
        """Cron job to check VAT thresholds for all companies based on 9-month rolling period"""
        _logger.info("Starting VAT threshold cron check (9-month rolling period)...")
        
        # Get all active companies
        companies = self.env['res.company'].search([])
        
        for company in companies:
            # Check if record exists for this company
            threshold_record = self.search([('company_id', '=', company.id)], limit=1)
            
            if not threshold_record:
                # Create new record
                threshold_record = self.create({'company_id': company.id})
            
            # Update the record - compute all fields including rolling period
            threshold_record._refresh_threshold_values()
            
            # Send email if needed and not already sent
            if threshold_record.threshold_action != 'none' and not threshold_record.email_sent:
                threshold_record.send_threshold_notification()
        
        _logger.info("VAT threshold cron check completed (9-month rolling period).")

    def action_verify_vat(self):
        """Mark VAT as verified"""
        self._ensure_system_user()
        self.write({'vat_verified': True})
        return {
            'type': 'ir.actions.client',
            'tag': 'reload',
        }

    def action_reset_email(self):
        """Reset email sent status to allow re-sending"""
        self._ensure_system_user()
        self.write({
            'email_sent': False,
            'email_sent_date': False
        })
        return {
            'type': 'ir.actions.client',
            'tag': 'reload',
        }

    def action_manual_update_rolling_period(self):
        """Manual refresh of rolling VAT values - Admin only"""
        self._ensure_system_user()
        self._refresh_threshold_values()
        return {
            'type': 'ir.actions.client',
            'tag': 'reload',
        }

    def action_create_missing_company_records(self):
        """Create VAT threshold records for all companies that don't have one - Admin only"""
        self._ensure_system_user()
        companies = self.env['res.company'].search([])
        existing_thresholds = self.sudo().search([])
        existing_company_ids = existing_thresholds.mapped('company_id.id')
        
        created_count = 0
        for company in companies:
            if company.id not in existing_company_ids:
                self.create({'company_id': company.id})
                created_count += 1
                _logger.info("Created VAT threshold record for company: %s", company.name)
        
        # Force recompute all records
        all_records = self.search([])
        all_records._refresh_threshold_values()
        
        return {
            'type': 'ir.actions.client',
            'tag': 'reload',
        }

    def action_recompute_all_records(self):
        """Force recompute on selected records, or all records if nothing is selected."""
        self._ensure_system_user()
        records = self if self else self.search([])
        records._refresh_threshold_values()
        
        return {
            'type': 'ir.actions.client',
            'tag': 'reload',
        }

    def action_batch_mark_verified(self):
        """Mark selected records as VAT verified - Admin only"""
        self._ensure_system_user()
        for record in self:
            record.vat_verified = True
        
        return {
            'type': 'ir.actions.client',
            'tag': 'reload',
        }

    def action_batch_mark_unverified(self):
        """Mark selected records as VAT unverified - Admin only"""
        self._ensure_system_user()
        for record in self:
            record.vat_verified = False
        
        return {
            'type': 'ir.actions.client',
            'tag': 'reload',
        }

    def action_open_settings(self):
        """Open VAT Threshold Settings to configure email recipients"""
        self._ensure_system_user()
        return {
            'type': 'ir.actions.act_window',
            'name': 'VAT Threshold Settings',
            'res_model': 'vat.threshold.config',
            'view_mode': 'form',
            'target': 'new',
            'view_id': self.env.ref('vat_threshold.vat_threshold_config_view_form').id,
        }

    def action_test_send_email(self):
        """Test button to send a test email to configured recipients"""
        self._ensure_system_user()
        # Get recipients from system parameter
        recipients_str = self.env['ir.config_parameter'].sudo().get_param('vat_threshold.daily_report_recipients', '')
        
        if not recipients_str:
            return {
                'type': 'ir.actions.client',
                'tag': 'display_notification',
                'params': {
                    'title': 'Error',
                    'message': 'No email recipients configured! Please configure in Settings first.',
                    'type': 'danger',
                    'sticky': False,
                }
            }
        
        # Parse email addresses
        emails = [e.strip() for e in recipients_str.split(',') if e.strip()]
        
        if not emails:
            return {
                'type': 'ir.actions.client',
                'tag': 'display_notification',
                'params': {
                    'title': 'Error',
                    'message': 'No valid email addresses found!',
                    'type': 'danger',
                    'sticky': False,
                }
            }
        
        # Get all threshold records for the test email
        all_records = self.search([])
        register_list = all_records.filtered(lambda r: r.threshold_action == 'register')
        cancel_list = all_records.filtered(lambda r: r.threshold_action == 'cancel')
        no_action_list = all_records.filtered(lambda r: r.threshold_action == 'none')
        
        # Prepare email body
        email_body = self._prepare_daily_report_body(register_list, cancel_list, no_action_list)
        
        # Send test email
        try:
            mail_values = {
                'subject': f'[TEST] Daily VAT Threshold Report - {fields.Date.today()}',
                'body_html': email_body,
                'email_to': ', '.join(emails),
                'email_from': self.env.user.email or 'noreply@example.com',
            }
            
            self.env['mail.mail'].create(mail_values).send()
            
            return {
                'type': 'ir.actions.client',
                'tag': 'display_notification',
                'params': {
                    'title': 'Success',
                    'message': f'Test email sent successfully to: {", ".join(emails)}',
                    'type': 'success',
                    'sticky': False,
                }
            }
            
        except Exception as e:
            _logger.error("Failed to send test email: %s", str(e))
            return {
                'type': 'ir.actions.client',
                'tag': 'display_notification',
                'params': {
                    'title': 'Error',
                    'message': f'Failed to send email: {str(e)}',
                    'type': 'danger',
                    'sticky': False,
                }
            }

    @api.model
    def _cron_send_daily_report(self):
        """Cron job to send daily consolidated email report at 7:05 AM
        
        Automatically sends email if ANY company has threshold_action != 'none'
        Uses email recipients from Settings (vat_threshold.daily_report_recipients)
        """
        _logger.info("Starting daily VAT threshold email report...")
        
        # Get all threshold records
        all_records = self.search([])
        
        # Separate into two lists
        register_list = all_records.filtered(lambda r: r.threshold_action == 'register')
        cancel_list = all_records.filtered(lambda r: r.threshold_action == 'cancel')
        no_action_list = all_records.filtered(lambda r: r.threshold_action == 'none')
        
        # Check if ANY action is detected - if not, skip sending email
        if not register_list and not cancel_list:
            _logger.info("No threshold actions detected - skipping email report")
            return
        
        # Get recipients from system parameter
        recipients_str = self.env['ir.config_parameter'].sudo().get_param('vat_threshold.daily_report_recipients', '')
        
        if not recipients_str:
            _logger.info("No recipients configured in Settings - skipping email report")
            return
        
        # Parse email addresses
        emails = [e.strip() for e in recipients_str.split(',') if e.strip()]
        
        if not emails:
            _logger.info("No valid email addresses found")
            return
        
        # Prepare email body
        email_body = self._prepare_daily_report_body(register_list, cancel_list, no_action_list)
        
        # Send email
        try:
            mail_values = {
                'subject': f'Daily VAT Threshold Report - {fields.Date.today()}',
                'body_html': email_body,
                'email_to': ', '.join(emails),
                'email_from': self.env.user.email or 'noreply@example.com',
            }
            
            self.env['mail.mail'].create(mail_values).send()
            _logger.info("Daily VAT threshold report sent to: %s", ', '.join(emails))
            
        except Exception as e:
            _logger.error("Failed to send daily VAT report: %s", str(e))

    def _prepare_daily_report_body(self, register_list, cancel_list, no_action_list):
        """Prepare HTML body for daily consolidated report"""
        
        def format_currency(amount):
            return f"{amount:,.2f} AED"
        
        # Build register list HTML
        register_html = ""
        if register_list:
            register_html = """
            <h3 style="color: #d9534f;">📋 LIST 1: COMPANIES REQUIRING VAT REGISTRATION</h3>
            <table style="border-collapse: collapse; width: 100%; margin-bottom: 20px;">
                <tr style="background-color: #f5f5f5;">
                    <th style="padding: 10px; border: 1px solid #ddd; text-align: left;">Company</th>
                    <th style="padding: 10px; border: 1px solid #ddd; text-align: right;">Rolling Sales (9 Months)</th>
                    <th style="padding: 10px; border: 1px solid #ddd; text-align: center;">Status</th>
                </tr>
            """
            for rec in register_list:
                register_html += f"""
                <tr>
                    <td style="padding: 10px; border: 1px solid #ddd;">{rec.company_id.name}</td>
                    <td style="padding: 10px; border: 1px solid #ddd; text-align: right;">{format_currency(rec.rolling_total_sales)}</td>
                    <td style="padding: 10px; border: 1px solid #ddd; text-align: center; color: #d9534f;">⚠️ Register VAT</td>
                </tr>
                """
            register_html += f"""
            </table>
            <p><strong>Total: {len(register_list)} companies require VAT registration</strong></p>
            """
        else:
            register_html = """
            <h3 style="color: #5cb85c;">📋 LIST 1: COMPANIES REQUIRING VAT REGISTRATION</h3>
            <p style="color: #5cb85c; padding: 10px; background-color: #dff0d8; border-radius: 5px;">✅ No companies require VAT registration at this time.</p>
            """
        
        # Build cancel list HTML
        cancel_html = ""
        if cancel_list:
            cancel_html = """
            <h3 style="color: #f0ad4e;">📋 LIST 2: COMPANIES REQUIRING VAT CANCELLATION</h3>
            <table style="border-collapse: collapse; width: 100%; margin-bottom: 20px;">
                <tr style="background-color: #f5f5f5;">
                    <th style="padding: 10px; border: 1px solid #ddd; text-align: left;">Company</th>
                    <th style="padding: 10px; border: 1px solid #ddd; text-align: right;">Rolling Sales (9 Months)</th>
                    <th style="padding: 10px; border: 1px solid #ddd; text-align: center;">Status</th>
                </tr>
            """
            for rec in cancel_list:
                cancel_html += f"""
                <tr>
                    <td style="padding: 10px; border: 1px solid #ddd;">{rec.company_id.name}</td>
                    <td style="padding: 10px; border: 1px solid #ddd; text-align: right;">{format_currency(rec.rolling_total_sales)}</td>
                    <td style="padding: 10px; border: 1px solid #ddd; text-align: center; color: #f0ad4e;">⚠️ Cancel VAT</td>
                </tr>
                """
            cancel_html += f"""
            </table>
            <p><strong>Total: {len(cancel_list)} companies eligible for VAT cancellation</strong></p>
            """
        else:
            cancel_html = """
            <h3 style="color: #5cb85c;">📋 LIST 2: COMPANIES REQUIRING VAT CANCELLATION</h3>
            <p style="color: #5cb85c; padding: 10px; background-color: #dff0d8; border-radius: 5px;">✅ No companies require VAT cancellation at this time.</p>
            """
        
        # Summary
        total_checked = len(register_list) + len(cancel_list) + len(no_action_list)
        summary_html = f"""
        <h3 style="color: #333;">📊 SUMMARY</h3>
        <table style="border-collapse: collapse; width: 100%; max-width: 400px; margin-top: 15px;">
            <tr>
                <td style="padding: 8px; border: 1px solid #ddd; background-color: #f5f5f5;"><strong>Total Companies Checked:</strong></td>
                <td style="padding: 8px; border: 1px solid #ddd;">{total_checked}</td>
            </tr>
            <tr>
                <td style="padding: 8px; border: 1px solid #ddd; background-color: #fcf8e3;"><strong>Require Registration:</strong></td>
                <td style="padding: 8px; border: 1px solid #ddd; color: #d9534f;">{len(register_list)}</td>
            </tr>
            <tr>
                <td style="padding: 8px; border: 1px solid #ddd; background-color: #fcf8e3;"><strong>Require Cancellation:</strong></td>
                <td style="padding: 8px; border: 1px solid #ddd; color: #f0ad4e;">{len(cancel_list)}</td>
            </tr>
            <tr>
                <td style="padding: 8px; border: 1px solid #ddd; background-color: #dff0d8;"><strong>No Action Required:</strong></td>
                <td style="padding: 8px; border: 1px solid #ddd; color: #5cb85c;">{len(no_action_list)}</td>
            </tr>
        </table>
        """
        
        # Full email body
        body = f"""
        <div style="font-family: Arial, sans-serif; padding: 20px; max-width: 800px;">
            <h2 style="color: #333; border-bottom: 2px solid #875A7B; padding-bottom: 10px;">
                Daily VAT Threshold Report - {fields.Date.today()}
            </h2>
            
            <p>Dear Team,</p>
            <p>Please find below the daily VAT threshold status report based on the 9-month rolling period calculation.</p>
            
            <hr style="border: 1px solid #eee; margin: 20px 0;"/>
            
            {register_html}
            
            <hr style="border: 1px solid #eee; margin: 20px 0;"/>
            
            {cancel_html}
            
            <hr style="border: 1px solid #eee; margin: 20px 0;"/>
            
            {summary_html}
            
            <br/>
            <p style="color: #666; font-size: 12px;">
                <strong>Threshold Limits:</strong><br/>
                • Registration Required: Sales > 375,000 AED (9-month rolling)<br/>
                • Cancellation Possible: Sales < 125,000 AED (9-month rolling)
            </p>
            
            <br/>
            <p>Best regards,</p>
            <p><strong>System Administrator</strong></p>
        </div>
        """
        
        return body
