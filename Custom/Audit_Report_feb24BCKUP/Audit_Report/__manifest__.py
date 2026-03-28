{
    'name': 'Audit Report',
    'version': '19.0.1.0.0',
    'category': 'Accounting',
    'summary': 'Generate, edit, and store audit report PDF revisions by company',
    'description': """
        This module provides a wizard to generate audit reports and manage
        saved, versioned, company-scoped editable PDF revisions.
    """,
    'author': 'Your Company',
    'license': 'LGPL-3',
    'depends': [
        'account_reports',
        'company_exte',
        'account_cash_flow_config',
    ],
    'external_dependencies': {
        'python': ['weasyprint'],
    },
    'data': [
        'security/ir.model.access.csv',
        'security/audit_report_security.xml',
        'views/views.xml',
    ],
    'demo': [],
    'installable': True,
    'auto_install': False,
    'application': False,
}
