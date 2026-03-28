{
    'name': 'VAT Threshold Management',
    'version': '19.0.1.1.0',
    'category': 'Accounting',
    'summary': 'Manage VAT threshold checks and notifications',
    'description': """
        This module tracks company revenue and sends email notifications
        based on VAT threshold conditions:
        - If VAT Registered and Sales < 125,000: Email to cancel VAT
        - If Not VAT Registered and Sales > 375,000: Email to register VAT
        
        Features:
        - 9-month rolling period calculation
        - Auto-create threshold records for new companies
        - Admin-only manual controls
    """,
    'author': 'Your Company',
    'license': 'LGPL-3',
    'depends': [
        'base',
        'account',
        'mail',
        'company_exte',
        'Audit_Report',
    ],
    'data': [
        'security/vat_threshold_security.xml',
        'security/ir.model.access.csv',
        'views/vat_threshold_view.xml',
        'views/vat_threshold_config_view.xml',
        'data/email_template.xml',
        'data/cron.xml',
    ],
    'post_init_hook': 'post_init_setup',
    'installable': True,
    'auto_install': False,
    'application': True,
}
