from odoo import fields, models


class IncentiveSegment(models.Model):
    _name = "incentive.segment"
    _description = "Incentive Business Segment"
    _order = "sequence, name"

    name = fields.Char(required=True)
    code = fields.Char(required=True, copy=False)
    sequence = fields.Integer(default=10)
    active = fields.Boolean(default=True)
    basis_type = fields.Selection(
        [
            ("commission_net", "Commission Net of Taxes"),
            ("profit_excl_gst", "Profit Excluding GST"),
            ("gross_salary", "Gross Salary"),
            ("commission_flat", "Fixed Commission Rule"),
        ],
        required=True,
        default="commission_net",
    )
    company_id = fields.Many2one("res.company", default=lambda self: self.env.company, required=True)
    gp_minimum = fields.Float(string="Default Minimum GP %", default=15.0)
    use_receipt_exchange = fields.Boolean(default=False)
    description = fields.Text()

    _sql_constraints = [
        ("incentive_segment_code_uniq", "unique(code, company_id)", "Segment code must be unique per company."),
    ]
