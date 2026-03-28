{
    'name': 'Custom Module Permissions - Company Extra Info',
    'version': '19.0.1.0.0',
    'summary': 'Company extra info visibility bridge for custom module permissions',
    'author': 'TI Associates',
    'license': 'LGPL-3',
    'category': 'Tools',
    'depends': [
        'custom_module_permissions',
        'company_exte',
    ],
    'data': [
        'security/company_exte_groups.xml',
        'views/company_exte_visibility.xml',
    ],
    'post_init_hook': 'post_init_hook',
    'installable': False,
    'application': False,
}
