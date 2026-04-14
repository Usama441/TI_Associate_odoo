{
    "name": "Target-Based Incentive & Commission Management",
    "version": "19.0.1.0.0",
    "summary": "Manage target-based incentive, commission, settlements, and exceptions.",
    "category": "Human Resources",
    "author": "OpenAI",
    "license": "LGPL-3",
    "depends": ["base", "mail", "hr", "account"],
    "data": [
        "security/security.xml",
        "security/ir.model.access.csv",
        "data/sequence_data.xml",
        "data/master_data.xml",
        "views/hr_employee_views.xml",
        "views/policy_views.xml",
        "views/period_views.xml",
        "views/target_views.xml",
        "views/source_views.xml",
        "views/exception_views.xml",
        "report/incentive_reports.xml",
        "views/settlement_views.xml"
    ],
    "application": True,
    "installable": True
}
