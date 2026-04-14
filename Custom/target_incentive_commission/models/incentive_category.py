from odoo import fields, models


class IncentiveEmployeeCategory(models.Model):
    _name = "incentive.employee.category"
    _description = "Incentive Employee Category"
    _order = "sequence, name"

    name = fields.Char(required=True)
    code = fields.Char(required=True, copy=False)
    sequence = fields.Integer(default=10)
    active = fields.Boolean(default=True)
    description = fields.Text()

    _sql_constraints = [
        ("incentive_employee_category_code_uniq", "unique(code)", "Category code must be unique."),
    ]
