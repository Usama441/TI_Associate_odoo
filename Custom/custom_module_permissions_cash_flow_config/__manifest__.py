{
    'name': 'Custom Module Permissions - Cash Flow Config',
    'version': '19.0.1.0.0',
    'summary': 'Cash flow config visibility bridge for custom module permissions',
    'author': 'TI Associates',
    'license': 'LGPL-3',
    'category': 'Tools',
    'depends': [
        'custom_module_permissions',
        'account_cash_flow_config',
    ],
    'data': [
        'security/cash_flow_config_groups.xml',
        'views/account_cash_flow_visibility.xml',
    ],
    'post_init_hook': 'post_init_hook',
    'installable': False,
    'application': False,
}
