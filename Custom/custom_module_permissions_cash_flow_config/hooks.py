from odoo import Command


def post_init_hook(env):
    group = env.ref('custom_module_permissions_cash_flow_config.group_custom_module_cash_flow_config', raise_if_not_found=False)
    if not group:
        return
    internal_users = env['res.users'].with_context(active_test=False).search([
        ('share', '=', False),
    ])
    if not internal_users:
        return
    internal_users.write({'group_ids': [Command.link(group.id)]})
