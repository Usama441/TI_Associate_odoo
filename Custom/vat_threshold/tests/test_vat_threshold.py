from odoo import Command, fields
from odoo.exceptions import AccessError
from odoo.tests import TransactionCase, tagged


@tagged('post_install', '-at_install')
class TestVatThreshold(TransactionCase):
    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        cls.group_user = cls.env.ref('base.group_user')
        cls.group_system = cls.env.ref('base.group_system')

    def _make_company(self, name):
        company = self.env['res.company'].create({'name': name})
        threshold = self.env['vat.threshold'].search([('company_id', '=', company.id)], limit=1)
        self.assertTrue(threshold, 'VAT threshold record should be auto-created for new companies')
        return company

    def _make_internal_user(self, company, prefix='vat_threshold_user', groups=None):
        login = f'{prefix}_{company.id}'
        return self.env['res.users'].with_context(no_reset_password=True).create({
            'name': login,
            'login': login,
            'email': f'{login}@example.com',
            'company_id': company.id,
            'company_ids': [Command.set([company.id])],
            'group_ids': [Command.set(groups or [self.group_user.id])],
        })

    def _make_journal(self, company):
        journal = self.env['account.journal'].search([
            ('company_id', '=', company.id),
            ('type', '=', 'general'),
        ], limit=1)
        if journal:
            return journal
        return self.env['account.journal'].create({
            'name': f'{company.name} Misc Journal',
            'code': f'VT{company.id % 1000:03d}',
            'type': 'general',
            'company_id': company.id,
        })

    def _make_account(self, company, code, name, account_type='income'):
        return self.env['account.account'].create({
            'name': name,
            'code': code,
            'account_type': account_type,
            'company_ids': [Command.set([company.id])],
        })

    def _post_entry(self, company, journal, credit_lines, debit_account):
        total_credit = sum(amount for _, amount in credit_lines)
        line_commands = [
            Command.create({
                'name': account.display_name,
                'account_id': account.id,
                'credit': amount,
                'debit': 0.0,
            })
            for account, amount in credit_lines
        ]
        line_commands.append(Command.create({
            'name': 'Counterpart',
            'account_id': debit_account.id,
            'debit': total_credit,
            'credit': 0.0,
        }))
        move = self.env['account.move'].create({
            'journal_id': journal.id,
            'date': fields.Date.today(),
            'line_ids': line_commands,
        })
        move.action_post()
        return move

    def test_auto_creates_threshold_record_for_new_company(self):
        company = self._make_company('VAT Auto Create Co')
        threshold = self.env['vat.threshold'].search([('company_id', '=', company.id)], limit=1)

        self.assertTrue(threshold)
        self.assertEqual(threshold.company_id, company)

    def test_company_rule_limits_visibility_to_allowed_company(self):
        company_a = self._make_company('VAT Scope A')
        company_b = self._make_company('VAT Scope B')
        user = self._make_internal_user(company_a, prefix='vat_scope')

        records = self.env['vat.threshold'].with_user(user).search([])

        self.assertEqual(set(records.mapped('company_id.id')), {company_a.id})
        self.assertEqual(len(records), 1)
        self.assertFalse(records.filtered(lambda r: r.company_id.id == company_b.id))

    def test_admin_actions_are_restricted_to_system_users(self):
        company = self._make_company('VAT Admin Guard Co')
        record = self.env['vat.threshold'].search([('company_id', '=', company.id)], limit=1)
        user = self._make_internal_user(company, prefix='vat_guard')

        with self.assertRaises(AccessError):
            record.with_user(user).action_check_threshold()

        with self.assertRaises(AccessError):
            record.with_user(user).action_open_settings()

    def test_refresh_recomputes_sales_and_resets_email_flags(self):
        company = self._make_company('VAT Revenue Co')
        journal = self._make_journal(company)
        revenue_uae = self._make_account(company, '41010101', 'UAE Client Revenue')
        revenue_intl = self._make_account(company, '41010102', 'International Revenue')
        revenue_related = self._make_account(company, '41020101', 'Related Party Revenue')
        counter_account = self._make_account(company, '99999998', 'VAT Counterpart', account_type='asset_current')

        self._post_entry(
            company,
            journal,
            [
                (revenue_uae, 150000.0),
                (revenue_intl, 100000.0),
                (revenue_related, 200000.0),
            ],
            counter_account,
        )

        record = self.env['vat.threshold'].search([('company_id', '=', company.id)], limit=1)
        record._refresh_threshold_values()

        self.assertEqual(record.rolling_revenue_uae_clients, 150000.0)
        self.assertEqual(record.rolling_revenue_international_clients, 100000.0)
        self.assertEqual(record.rolling_revenue_related_party, 200000.0)
        self.assertEqual(record.rolling_total_sales, 450000.0)
        self.assertEqual(record.threshold_action, 'register')
        self.assertFalse(record.email_sent)
        self.assertFalse(record.email_sent_date)

        record.write({
            'email_sent': True,
            'email_sent_date': fields.Datetime.now(),
        })
        company.sudo().write({'vat_registration_number': 'AE123456789'})
        company.invalidate_recordset(['vat_registration_number'])

        record._refresh_threshold_values()

        self.assertTrue(record.is_vat_registered)
        self.assertEqual(record.threshold_action, 'none')
        self.assertFalse(record.email_sent)
        self.assertFalse(record.email_sent_date)

    def test_admin_only_actions_and_settings_are_bound_to_system_group(self):
        recompute_action = self.env.ref('vat_threshold.action_recompute_all_records')
        create_missing_action = self.env.ref('vat_threshold.action_create_missing_companies')
        batch_verify_action = self.env.ref('vat_threshold.action_batch_mark_verified')
        batch_unverify_action = self.env.ref('vat_threshold.action_batch_mark_unverified')
        settings_action = self.env.ref('vat_threshold.action_vat_threshold_config')

        self.assertEqual(set(recompute_action.group_ids.ids), {self.group_system.id})
        self.assertEqual(set(create_missing_action.group_ids.ids), {self.group_system.id})
        self.assertEqual(set(batch_verify_action.group_ids.ids), {self.group_system.id})
        self.assertEqual(set(batch_unverify_action.group_ids.ids), {self.group_system.id})
        self.assertEqual(set(settings_action.group_ids.ids), {self.group_system.id})
