from unittest.mock import patch
from odoo import fields
from odoo.exceptions import ValidationError
from odoo.tests import TransactionCase, tagged


@tagged('post_install', '-at_install')
class TestAuditReportFreeZoneDefaults(TransactionCase):
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
            'company_free_zone': 'Meydan Free Zone',
            'company_license_number': 'LIC-001',
            'trade_license_activities': 'Consulting services',
            'incorporation_date': fields.Date.to_date('2020-01-01'),
        }
        values.update(overrides)
        return self.env['audit.report'].create(values)

    def test_company_free_zone_location_defaults_cover_requested_mappings(self):
        expected_defaults = {
            'Dubai Integrated Economic Zones Authority': {
                'street': 'DSO-IFZA, IFZA Properties, Dubai Silicon Oasis',
                'city': 'Dubai',
            },
            'Dubai International Financial Centre': {
                'city': 'Dubai',
            },
            'Meydan Free Zone': {
                'street': 'Meydan Grandstand, 6th Floor, Meydan Road, Nad Al Sheba',
                'city': 'Dubai',
            },
            'Sharjah Media City': {
                'street': 'Sharjah Media City',
                'city': 'Sharjah',
            },
            'Creative Media City': {
                'street': 'Fujairah - Creative Tower, P.O.Box 4422',
                'city': 'Fujairah',
            },
            'Sharjah Publishing City': {
                'street': 'Business Centre, Sharjah Publishing City Freezone',
                'city': 'Sharjah',
            },
        }

        for free_zone, expected in expected_defaults.items():
            with self.subTest(free_zone=free_zone):
                self.assertEqual(
                    self.company._get_free_zone_location_defaults(free_zone),
                    expected,
                )

    def test_dmcc_companies_require_portal_account_before_render(self):
        wizard = self._create_wizard(signature_include_1=True)
        wizard.company_free_zone = 'Dubai Multi Commodities Centre Free Zone'
        wizard.portal_account_no = False

        with self.assertRaisesRegex(ValidationError, 'Portal Account'):
            wizard._validate_emphasis_options()

    def test_first_shareholder_signatory_is_required(self):
        wizard = self._create_wizard(signature_include_1=False)

        with self.assertRaisesRegex(ValidationError, 'SIGNATORY REQUIRED'):
            wizard._validate_emphasis_options()

    def test_two_year_reports_require_soce_label_date(self):
        wizard = self._create_wizard(
            audit_period_category='normal_2y',
            signature_include_1=True,
            soce_prior_opening_label_date=False,
        )

        with self.assertRaisesRegex(ValidationError, 'SOCE FIRST BALANCE DATE REQUIRED'):
            wizard._validate_emphasis_options()

    def test_company_free_zone_is_required_before_render(self):
        wizard = self._create_wizard(
            signature_include_1=True,
            company_free_zone=False,
        )

        with self.assertRaisesRegex(ValidationError, 'FREE ZONE REQUIRED'):
            wizard._validate_emphasis_options()

    def test_company_license_number_is_required_before_render(self):
        wizard = self._create_wizard(
            signature_include_1=True,
            company_license_number=False,
        )

        with self.assertRaisesRegex(ValidationError, 'COMPANY LICENSE NUMBER REQUIRED'):
            wizard._validate_emphasis_options()

    def test_trade_license_details_are_required_before_render(self):
        wizard = self._create_wizard(
            signature_include_1=True,
            trade_license_activities=False,
        )

        with self.assertRaisesRegex(ValidationError, 'TRADE LICENSE NUMBER REQUIRED'):
            wizard._validate_emphasis_options()

    def test_corporate_incorporation_details_are_required_before_render(self):
        wizard = self._create_wizard(
            signature_include_1=True,
            incorporation_date=False,
        )

        with self.assertRaisesRegex(ValidationError, 'CORPORATE INCORPORATION NUMBER REQUIRED'):
            wizard._validate_emphasis_options()

    def test_owner_current_account_zero_does_not_change_share_capital_status_on_validate(self):
        wizard = self._create_wizard(
            signature_include_1=True,
            share_capital_paid_status='unpaid',
            nationality_1='Pakistani',
        )
        wizard.shareholder_1 = 'Primary Shareholder'

        with patch.object(type(wizard), '_get_owner_current_account_equity_for_status', return_value=0.0):
            wizard._validate_emphasis_options()

        self.assertEqual(wizard.share_capital_paid_status, 'unpaid')

    def test_owner_current_account_zero_does_not_change_share_capital_status_on_onchange(self):
        wizard = self._create_wizard(share_capital_paid_status='unpaid')

        with patch.object(type(wizard), '_get_owner_current_account_equity_for_status', return_value=0.0):
            result = wizard._onchange_share_capital_paid_status()

        self.assertEqual(result, {})
        self.assertEqual(wizard.share_capital_paid_status, 'unpaid')

    def test_non_zero_owner_current_account_returns_share_capital_warning_on_onchange(self):
        wizard = self._create_wizard(
            signature_include_1=True,
            share_capital_paid_status='paid',
        )

        with patch.object(type(wizard), '_get_owner_current_account_equity_for_status', return_value=42.0):
            result = wizard._onchange_share_capital_paid_status()

        self.assertIn('warning', result)
        self.assertIn('Owner current account is not zero', result['warning']['message'])
        self.assertEqual(wizard.share_capital_paid_status, 'paid')

    def test_audit_report_onchange_company_free_zone_sets_address_and_city(self):
        wizard = self._create_wizard()

        wizard.company_free_zone = 'Meydan Free Zone'
        wizard._onchange_company_free_zone()

        self.assertEqual(
            wizard.company_street,
            'Meydan Grandstand, 6th Floor, Meydan Road, Nad Al Sheba',
        )
        self.assertEqual(wizard.company_city, 'Dubai')

    def test_audit_report_onchange_company_free_zone_sets_difc_regulation_and_city(self):
        wizard = self._create_wizard()

        wizard.company_free_zone = 'Dubai International Financial Centre'
        wizard._onchange_company_free_zone()

        self.assertEqual(wizard.company_city, 'Dubai')
        self.assertEqual(self.company.city, 'Dubai')
        self.assertEqual(
            wizard.implementing_regulations_freezone,
            'Dubai International Financial Centre Companies Law No. 5 of 2018',
        )
        self.assertEqual(
            self.company.implementing_regulations_freezone,
            'Dubai International Financial Centre Companies Law No. 5 of 2018',
        )

    def test_company_free_zone_location_defaults_infer_city_when_address_is_unknown(self):
        expected_defaults = {
            'Abu Dhabi Global Market': {'city': 'Abu Dhabi'},
            'Ajman Free zone': {'city': 'Ajman'},
            'Dubai Multi Commodities Centre Free Zone': {'city': 'Dubai'},
            'Department of Economic Development': {'city': 'Dubai'},
            'Department of Economy and Tourism': {'city': 'Dubai'},
            'Ras Al Khaimah Economic Zone': {'city': 'Ras Al Khaimah'},
            'Ras Al Khaimah International Corporate Centre': {'city': 'Ras Al Khaimah'},
            'Sharjah Research Technology and Innovation Park Free Zone Authority': {'city': 'Sharjah'},
        }

        for free_zone, expected in expected_defaults.items():
            with self.subTest(free_zone=free_zone):
                self.assertEqual(
                    self.company._get_free_zone_location_defaults(free_zone),
                    expected,
                )
