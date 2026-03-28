from odoo import Command
from odoo.addons.custom_module_permissions_company_exte.hooks import post_init_hook
from odoo.tests import TransactionCase, tagged


@tagged('post_install', '-at_install')
class TestCompanyExtePermissions(TransactionCase):
    def test_group_and_visibility_views(self):
        group = self.env.ref('custom_module_permissions_company_exte.group_custom_module_company_other_info')
        self.assertTrue(group)
        self.assertEqual(
            group.privilege_id.category_id,
            self.env.ref('custom_module_permissions.module_category_custom_module_access'),
        )
        self.assertIn(self.env.ref('base.group_user'), group.implied_ids)

        view_a = self.env.ref('custom_module_permissions_company_exte.view_company_form_other_information_visibility')
        self.assertIn('custom_module_permissions_company_exte.group_custom_module_company_other_info', view_a.arch_db)

        view_b = self.env.ref('custom_module_permissions_company_exte.view_company_form_shareholders_visibility')
        self.assertIn('custom_module_permissions_company_exte.group_custom_module_company_other_info', view_b.arch_db)

    def test_post_init_grants_group_to_internal_users(self):
        user = self.env['res.users'].with_context(no_reset_password=True).create({
            'name': 'bridge_company_user',
            'login': 'bridge_company_user',
            'company_id': self.env.company.id,
            'company_ids': [Command.set(self.env.company.ids)],
            'group_ids': [Command.set([self.env.ref('base.group_user').id])],
        })
        group = self.env.ref('custom_module_permissions_company_exte.group_custom_module_company_other_info')
        self.assertNotIn(group, user.group_ids)
        post_init_hook(self.env)
        user.invalidate_recordset(['group_ids'])
        self.assertIn(group, user.group_ids)
