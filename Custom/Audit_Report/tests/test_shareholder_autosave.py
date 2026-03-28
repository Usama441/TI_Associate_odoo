from odoo import fields
from odoo.tests import TransactionCase, tagged


@tagged('post_install', '-at_install')
class TestAuditReportShareholderAutosave(TransactionCase):
    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        cls.company = cls.env.company

    def _create_wizard(self, **overrides):
        values = {
            'company_id': self.company.id,
            'date_start': fields.Date.to_date('2024-01-01'),
            'date_end': fields.Date.to_date('2024-12-31'),
            'report_type': 'period',
            'audit_period_category': 'normal_1y',
            'signature_date_mode': 'today',
            'use_previous_settings': False,
        }
        values.update(overrides)
        return self.env['audit.report'].create(values)

    def test_onchange_shareholder_preferences_autosaves_previous_settings(self):
        wizard = self._create_wizard(
            share_capital_paid_status='paid',
            signature_include_2=False,
            director_include_2=False,
        )

        wizard.share_capital_paid_status = 'unpaid'
        wizard.signature_include_2 = True
        wizard.director_include_2 = True
        wizard._onchange_autosave_shareholder_preferences()

        previous = wizard._get_previous_settings()
        self.assertEqual(previous.get('share_capital_paid_status'), 'unpaid')
        self.assertTrue(previous.get('signature_include_2'))
        self.assertTrue(previous.get('director_include_2'))

    def test_onchange_shareholder_preferences_syncs_company_share(self):
        wizard = self._create_wizard(share_capital_paid_status='paid')

        wizard.share_capital_paid_status = 'unpaid'
        wizard._onchange_autosave_shareholder_preferences()

        self.company.invalidate_recordset(['company_share'])
        self.assertEqual(self.company.company_share, 'unpaid')
