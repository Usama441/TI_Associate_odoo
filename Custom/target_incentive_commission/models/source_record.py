from datetime import date

from odoo import api, fields, models
from odoo.exceptions import ValidationError


class IncentiveSourceRecord(models.Model):
    _name = "incentive.source.record"
    _description = "Incentive Source Record"
    _inherit = ["mail.thread", "mail.activity.mixin"]
    _order = "receipt_date desc, id desc"

    name = fields.Char(required=True, tracking=True)
    active = fields.Boolean(default=True)
    company_id = fields.Many2one("res.company", default=lambda self: self.env.company, required=True)
    company_currency_id = fields.Many2one("res.currency", related="company_id.currency_id", store=True, readonly=True)
    review_period_id = fields.Many2one("incentive.review.period", tracking=True)
    employee_id = fields.Many2one("hr.employee", required=True, tracking=True)
    category_id = fields.Many2one("incentive.employee.category", related="employee_id.incentive_category_id", store=True, readonly=True)
    segment_id = fields.Many2one("incentive.segment", required=True, tracking=True)
    principal_partner_id = fields.Many2one("res.partner", string="Principal")
    reference = fields.Char(help="Order / invoice / receipt reference")
    receipt_date = fields.Date(required=True, tracking=True)
    currency_id = fields.Many2one("res.currency", required=True, default=lambda self: self.env.company.currency_id.id, tracking=True)
    exchange_rate = fields.Float(default=1.0, digits=(16, 6), tracking=True)

    commission_amount = fields.Monetary(currency_field="currency_id")
    commission_tax_amount = fields.Monetary(currency_field="currency_id")
    commission_net_amount = fields.Monetary(currency_field="currency_id", compute="_compute_commission_net_amount", store=True)
    sales_amount = fields.Monetary(currency_field="currency_id")
    profit_excl_gst = fields.Monetary(currency_field="currency_id")
    gross_salary_amount = fields.Monetary(currency_field="company_currency_id", compute="_compute_gross_salary_amount", store=True)
    gp_percent = fields.Float(string="GP %", tracking=True)
    gp_threshold = fields.Float(compute="_compute_gp_threshold", store=True)
    gp_threshold_passed = fields.Boolean(compute="_compute_eligibility_flags", store=True)
    exception_allowed = fields.Boolean(compute="_compute_gp_threshold", store=True)
    exception_request_ids = fields.One2many("incentive.exception.request", "source_record_id")
    approved_exception_id = fields.Many2one("incentive.exception.request", compute="_compute_approved_exception", store=True)
    has_approved_exception = fields.Boolean(compute="_compute_approved_exception", store=True)

    assessed_basis_amount = fields.Monetary(currency_field="currency_id", compute="_compute_assessed_basis_amount", store=True)
    eligible = fields.Boolean(compute="_compute_eligibility_flags", store=True)
    eligibility_status = fields.Selection(
        [
            ("eligible", "Eligible"),
            ("blocked_gp", "Blocked: GP Below Threshold"),
            ("blocked_missing_policy", "Blocked: Missing Policy"),
            ("blocked_other", "Blocked: Other"),
        ],
        compute="_compute_eligibility_flags",
        store=True,
    )

    new_principal_flag = fields.Boolean(tracking=True)
    order_value_for_qualification = fields.Monetary(currency_field="currency_id")
    qualifies_value_threshold = fields.Boolean(compute="_compute_value_threshold", store=True)
    qualifying_order_count = fields.Integer(compute="_compute_qualifying_order_count")
    special_reward_eligible = fields.Boolean(compute="_compute_special_reward_eligible", search="_search_special_reward_eligible")

    state = fields.Selection([("draft", "Draft"), ("confirmed", "Confirmed"), ("cancelled", "Cancelled")], default="draft", tracking=True)
    note = fields.Text()
    policy_id = fields.Many2one("incentive.policy", compute="_compute_policy_id", store=True)

    @api.depends("commission_amount", "commission_tax_amount")
    def _compute_commission_net_amount(self):
        for rec in self:
            rec.commission_net_amount = (rec.commission_amount or 0.0) - (rec.commission_tax_amount or 0.0)

    @api.depends("employee_id", "employee_id.gross_salary_amount", "employee_id.salary_source_method")
    def _compute_gross_salary_amount(self):
        for rec in self:
            rec.gross_salary_amount = rec.employee_id._get_incentive_gross_salary() if rec.employee_id else 0.0

    @api.depends("segment_id", "category_id", "company_id")
    def _compute_policy_id(self):
        for rec in self:
            rec.policy_id = self.env["incentive.policy"].search(
                [
                    ("company_id", "=", rec.company_id.id),
                    ("category_id", "=", rec.category_id.id),
                    ("segment_id", "=", rec.segment_id.id),
                    ("active", "=", True),
                ],
                limit=1,
            )

    @api.depends("segment_id", "policy_id")
    def _compute_gp_threshold(self):
        for rec in self:
            rec.gp_threshold = rec.policy_id.gp_minimum or rec.segment_id.gp_minimum or 15.0
            rec.exception_allowed = bool(rec.policy_id.allow_gp_exception) if rec.policy_id else True

    @api.depends("exception_request_ids.state")
    def _compute_approved_exception(self):
        for rec in self:
            approved = rec.exception_request_ids.filtered(lambda x: x.state == "approved")[:1]
            rec.approved_exception_id = approved.id if approved else False
            rec.has_approved_exception = bool(approved)

    @api.depends("segment_id", "commission_net_amount", "profit_excl_gst", "gross_salary_amount")
    def _compute_assessed_basis_amount(self):
        for rec in self:
            basis = 0.0
            if rec.segment_id.basis_type == "commission_net":
                basis = rec.commission_net_amount
            elif rec.segment_id.basis_type == "profit_excl_gst":
                basis = rec.profit_excl_gst
            elif rec.segment_id.basis_type == "gross_salary":
                basis = rec.gross_salary_amount
            elif rec.segment_id.basis_type == "commission_flat":
                basis = rec.commission_net_amount
            rec.assessed_basis_amount = basis

    @api.depends("gp_percent", "gp_threshold", "policy_id", "has_approved_exception", "segment_id")
    def _compute_eligibility_flags(self):
        for rec in self:
            if not rec.policy_id and rec.segment_id.basis_type != "gross_salary":
                rec.eligible = False
                rec.gp_threshold_passed = False
                rec.eligibility_status = "blocked_missing_policy"
                continue

            gp_pass = rec.gp_percent >= rec.gp_threshold or rec.has_approved_exception
            rec.gp_threshold_passed = gp_pass
            if rec.segment_id.basis_type in ("commission_net", "profit_excl_gst", "commission_flat"):
                if not gp_pass:
                    rec.eligible = False
                    rec.eligibility_status = "blocked_gp"
                else:
                    rec.eligible = True
                    rec.eligibility_status = "eligible"
            else:
                rec.eligible = True
                rec.eligibility_status = "eligible"

    @api.depends("currency_id", "order_value_for_qualification", "company_currency_id")
    def _compute_value_threshold(self):
        usd = self.env.ref("base.USD", raise_if_not_found=False)
        for rec in self:
            qualifies = False
            if rec.currency_id and usd and rec.currency_id == usd:
                qualifies = (rec.order_value_for_qualification or 0.0) >= 50000.0
            elif rec.currency_id and rec.company_currency_id and rec.currency_id == rec.company_currency_id:
                qualifies = (rec.order_value_for_qualification or 0.0) >= 10000000.0
            rec.qualifies_value_threshold = qualifies

    @api.depends("new_principal_flag", "principal_partner_id", "receipt_date", "employee_id", "qualifies_value_threshold", "state")
    def _compute_qualifying_order_count(self):
        for rec in self:
            rec.qualifying_order_count = 0
            if not (rec.new_principal_flag and rec.principal_partner_id and rec.receipt_date):
                continue
            year_start = date(rec.receipt_date.year, 1, 1)
            year_end = date(rec.receipt_date.year, 12, 31)
            count = self.search_count(
                [
                    ("id", "!=", rec.id),
                    ("employee_id", "=", rec.employee_id.id),
                    ("principal_partner_id", "=", rec.principal_partner_id.id),
                    ("new_principal_flag", "=", True),
                    ("receipt_date", ">=", year_start),
                    ("receipt_date", "<=", year_end),
                    ("qualifies_value_threshold", "=", True),
                    ("state", "!=", "cancelled"),
                ]
            )
            rec.qualifying_order_count = count + (1 if rec.qualifies_value_threshold else 0)

    @api.depends("new_principal_flag", "qualifies_value_threshold", "qualifying_order_count")
    def _compute_special_reward_eligible(self):
        for rec in self:
            rec.special_reward_eligible = bool(rec.new_principal_flag and rec.qualifies_value_threshold and rec.qualifying_order_count in (2, 3))

    @api.model
    def _search_special_reward_eligible(self, operator, value):
        """Allow search views to filter this computed, non-stored field."""
        supported = {"=", "==", "!=", "<>", "in", "not in"}
        if operator not in supported:
            return [("id", "=", 0)]

        if operator in ("=", "=="):
            expected = bool(value)
        elif operator in ("!=", "<>"):
            expected = not bool(value)
        elif operator == "in":
            bool_values = {bool(v) for v in (value or [])}
            if bool_values == {True}:
                expected = True
            elif bool_values == {False}:
                expected = False
            else:
                return []
        else:  # operator == "not in"
            bool_values = {bool(v) for v in (value or [])}
            if bool_values == {True}:
                expected = False
            elif bool_values == {False}:
                expected = True
            else:
                return [("id", "=", 0)]

        candidate_domain = [
            ("new_principal_flag", "=", True),
            ("qualifies_value_threshold", "=", True),
            ("principal_partner_id", "!=", False),
            ("employee_id", "!=", False),
            ("receipt_date", "!=", False),
        ]
        matching_ids = self.search(candidate_domain).filtered("special_reward_eligible").ids
        return [("id", "in", matching_ids)] if expected else [("id", "not in", matching_ids)]

    @api.onchange("currency_id", "receipt_date")
    def _onchange_currency_receipt_date(self):
        if self.currency_id and self.company_currency_id:
            if self.currency_id == self.company_currency_id:
                self.exchange_rate = 1.0
            elif self.receipt_date:
                rate = self.currency_id._get_conversion_rate(self.currency_id, self.company_currency_id, self.company_id, self.receipt_date)
                self.exchange_rate = rate or 1.0

    @api.constrains("exchange_rate")
    def _check_exchange_rate(self):
        for rec in self:
            if rec.exchange_rate <= 0:
                raise ValidationError("Exchange rate must be greater than zero.")

    def action_confirm(self):
        self.write({"state": "confirmed"})

    def action_set_draft(self):
        self.write({"state": "draft"})

    def action_cancel(self):
        self.write({"state": "cancelled"})
