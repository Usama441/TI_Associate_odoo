from odoo import Command


def post_init_hook(env):
    group = env.ref('custom_module_permissions_audit_excel_export.group_custom_module_audit_excel_export', raise_if_not_found=False)
    if not group:
        return
    internal_users = env['res.users'].with_context(active_test=False).search([
        ('share', '=', False),
    ])
    if not internal_users:
        return
    internal_users.write({'group_ids': [Command.link(group.id)]})
