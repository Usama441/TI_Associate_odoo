from odoo import api, fields, models
from odoo.exceptions import ValidationError


class IncentivePolicy(models.Model):
    _name = "incentive.policy"
    _description = "Incentive Policy"
    _inherit = ["mail.thread", "mail.activity.mixin"]
    _order = "sequence, id"

    name = fields.Char(required=True, tracking=True)
    sequence = fields.Integer(default=10)
    active = fields.Boolean(default=True)
    company_id = fields.Many2one("res.company", default=lambda self: self.env.company, required=True, tracking=True)
    category_id = fields.Many2one("incentive.employee.category", required=True, tracking=True)
    segment_id = fields.Many2one("incentive.segment", required=True, tracking=True)
    approver_group_id = fields.Many2one("res.groups", string="Approver Group")
    settle_every_months = fields.Integer(default=6)
    gp_minimum = fields.Float(string="Minimum GP %", default=15.0)
    allow_gp_exception = fields.Boolean(default=True)
    allow_manual_override = fields.Boolean(default=True)
    bonus_if_all_segments_100 = fields.Boolean(string="Extra Gross Salary if All Segments >= 100%", default=True)
    special_reward_enabled = fields.Boolean(default=True)
    special_reward_percent = fields.Float(default=1.0)
    note = fields.Text()
    line_ids = fields.One2many("incentive.policy.line", "policy_id", copy=True)

    _sql_constraints = [
        ("incentive_policy_unique_scope", "unique(company_id, category_id, segment_id)", "Only one policy per company/category/segment is allowed."),
    ]

    @api.constrains("line_ids")
    def _check_lines(self):
        for rec in self:
            if not rec.line_ids:
                raise ValidationError("Each policy must contain at least one slab line.")


class IncentivePolicyLine(models.Model):
    _name = "incentive.policy.line"
    _description = "Incentive Policy Achievement Slab"
    _order = "policy_id, achievement_from, id"

    policy_id = fields.Many2one("incentive.policy", required=True, ondelete="cascade")
    sequence = fields.Integer(default=10)
    name = fields.Char(required=True)
    achievement_from = fields.Float(required=True)
    achievement_to = fields.Float(help="Inclusive upper bound. Leave empty for open ended.")
    calculation_type = fields.Selection(
        [("percent", "Percentage"), ("salary_multiplier", "Salary Multiplier"), ("fixed_amount", "Fixed Amount")],
        required=True,
        default="percent",
    )
    percentage = fields.Float()
    salary_multiplier = fields.Float()
    fixed_amount = fields.Float()
    description = fields.Char()

    @api.constrains("achievement_from", "achievement_to")
    def _check_range(self):
        for rec in self:
            if rec.achievement_to and rec.achievement_to < rec.achievement_from:
                raise ValidationError("Achievement upper bound must be greater than lower bound.")

    def matches(self, achievement_percent):
        self.ensure_one()
        lower_ok = achievement_percent >= self.achievement_from
        upper_ok = (not self.achievement_to) or achievement_percent <= self.achievement_to
        return lower_ok and upper_ok
