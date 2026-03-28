{
    'name': 'Custom Module Permissions - Audit Excel Export',
    'version': '19.0.1.0.0',
    'summary': 'Audit Excel Export visibility bridge for custom module permissions',
    'author': 'TI Associates',
    'license': 'LGPL-3',
    'category': 'Tools',
    'depends': [
        'custom_module_permissions',
        'audit_excel_export',
    ],
    'data': [
        'security/audit_excel_export_groups.xml',
        'views/audit_excel_export_visibility.xml',
    ],
    'post_init_hook': 'post_init_hook',
    'installable': False,
    'application': False,
}
