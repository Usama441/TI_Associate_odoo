{
    'name': 'Custom Module Permissions - Audit Report',
    'version': '19.0.1.0.0',
    'summary': 'Audit Report visibility bridge for custom module permissions',
    'author': 'TI Associates',
    'license': 'LGPL-3',
    'category': 'Tools',
    'depends': [
        'custom_module_permissions',
        'Audit_Report',
    ],
    'data': [
        'security/audit_report_groups.xml',
        'views/audit_report_visibility.xml',
    ],
    'post_init_hook': 'post_init_hook',
    'installable': False,
    'application': False,
}
