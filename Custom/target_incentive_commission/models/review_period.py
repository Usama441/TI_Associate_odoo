from odoo import api, fields, models
from odoo.exceptions import ValidationError


class IncentiveReviewPeriod(models.Model):
    _name = "incentive.review.period"
    _description = "Incentive Review / Settlement Period"
    _inherit = ["mail.thread", "mail.activity.mixin"]
    _order = "date_start desc, id desc"

    name = fields.Char(required=True, tracking=True)
    company_id = fields.Many2one("res.company", default=lambda self: self.env.company, required=True, tracking=True)
    period_type = fields.Selection(
        [
            ("quarterly", "Quarterly Review"),
            ("semiannual", "Semiannual Settlement"),
            ("annual", "Annual"),
        ],
        required=True,
        default="quarterly",
        tracking=True,
    )
    date_start = fields.Date(required=True, tracking=True)
    date_end = fields.Date(required=True, tracking=True)
    active = fields.Boolean(default=True)
    state = fields.Selection(
        [("draft", "Draft"), ("open", "Open"), ("closed", "Closed")],
        default="draft",
        tracking=True,
    )
    target_line_ids = fields.One2many("incentive.target.line", "review_period_id")
    source_record_ids = fields.One2many("incentive.source.record", "review_period_id")
    settlement_ids = fields.One2many("incentive.settlement", "period_id")
    note = fields.Text()

    @api.constrains("date_start", "date_end")
    def _check_dates(self):
        for rec in self:
            if rec.date_end < rec.date_start:
                raise ValidationError("End date must be greater than or equal to start date.")
