from odoo import api, fields, models
from odoo.exceptions import ValidationError


class IncentiveTargetLine(models.Model):
    _name = "incentive.target.line"
    _description = "Incentive Target / Achievement"
    _inherit = ["mail.thread", "mail.activity.mixin"]
    _order = "review_period_id desc, employee_id, segment_id"

    name = fields.Char(compute="_compute_name", store=True)
    active = fields.Boolean(default=True)
    company_id = fields.Many2one("res.company", default=lambda self: self.env.company, required=True)
    review_period_id = fields.Many2one("incentive.review.period", required=True, tracking=True)
    employee_id = fields.Many2one("hr.employee", required=True, tracking=True)
    department_id = fields.Many2one("hr.department", related="employee_id.department_id", store=True, readonly=True)
    category_id = fields.Many2one("incentive.employee.category", related="employee_id.incentive_category_id", store=True, readonly=True)
    segment_id = fields.Many2one("incentive.segment", required=True, tracking=True)
    target_amount = fields.Float(required=True, tracking=True)
    achievement_amount = fields.Float(default=0.0, tracking=True)
    achievement_percent = fields.Float(compute="_compute_achievement_percent", store=True)
    measure_note = fields.Char(string="Measure / Basis Note")
    is_manual = fields.Boolean(default=True)
    state = fields.Selection([("draft", "Draft"), ("confirmed", "Confirmed"), ("cancelled", "Cancelled")], default="draft", tracking=True)
    note = fields.Text()

    @api.depends("review_period_id", "employee_id", "segment_id")
    def _compute_name(self):
        for rec in self:
            rec.name = " / ".join([x for x in [rec.review_period_id.name, rec.employee_id.name, rec.segment_id.name] if x])

    @api.depends("target_amount", "achievement_amount")
    def _compute_achievement_percent(self):
        for rec in self:
            rec.achievement_percent = ((rec.achievement_amount or 0.0) / rec.target_amount * 100.0) if rec.target_amount else 0.0

    @api.constrains("target_amount")
    def _check_target_amount(self):
        for rec in self:
            if rec.target_amount < 0:
                raise ValidationError("Target amount cannot be negative.")

    def action_confirm(self):
        self.write({"state": "confirmed"})

    def action_set_draft(self):
        self.write({"state": "draft"})

    def action_cancel(self):
        self.write({"state": "cancelled"})
