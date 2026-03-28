from odoo import Command
from odoo.addons.custom_module_permissions_cash_flow_config.hooks import post_init_hook
from odoo.tests import TransactionCase, tagged


@tagged('post_install', '-at_install')
class TestCashFlowPermissions(TransactionCase):
    def test_group_and_visibility_view(self):
        group = self.env.ref('custom_module_permissions_cash_flow_config.group_custom_module_cash_flow_config')
        self.assertTrue(group)
        self.assertEqual(
            group.privilege_id.category_id,
            self.env.ref('custom_module_permissions.module_category_custom_module_access'),
        )
        self.assertIn(self.env.ref('base.group_user'), group.implied_ids)

        view = self.env.ref('custom_module_permissions_cash_flow_config.view_account_account_form_cash_flow_visibility')
        self.assertIn('custom_module_permissions_cash_flow_config.group_custom_module_cash_flow_config', view.arch_db)

    def test_post_init_grants_group_to_internal_users(self):
        user = self.env['res.users'].with_context(no_reset_password=True).create({
            'name': 'bridge_cash_user',
            'login': 'bridge_cash_user',
            'company_id': self.env.company.id,
            'company_ids': [Command.set(self.env.company.ids)],
            'group_ids': [Command.set([self.env.ref('base.group_user').id])],
        })
        group = self.env.ref('custom_module_permissions_cash_flow_config.group_custom_module_cash_flow_config')
        self.assertNotIn(group, user.group_ids)
        post_init_hook(self.env)
        user.invalidate_recordset(['group_ids'])
        self.assertIn(group, user.group_ids)
