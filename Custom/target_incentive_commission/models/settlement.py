from collections import defaultdict

from odoo import api, fields, models
from odoo.exceptions import UserError, ValidationError


class IncentiveSettlement(models.Model):
    _name = "incentive.settlement"
    _description = "Incentive Settlement"
    _inherit = ["mail.thread", "mail.activity.mixin"]
    _order = "period_id desc, id desc"

    name = fields.Char(required=True, copy=False, default="New", tracking=True)
    company_id = fields.Many2one("res.company", default=lambda self: self.env.company, required=True)
    period_id = fields.Many2one("incentive.review.period", required=True, domain=[("period_type", "=", "semiannual")], tracking=True)
    date_start = fields.Date(related="period_id.date_start", store=True, readonly=True)
    date_end = fields.Date(related="period_id.date_end", store=True, readonly=True)
    state = fields.Selection(
        [("draft", "Draft"), ("reviewed", "Reviewed"), ("approved", "Approved"), ("paid", "Paid"), ("cancelled", "Cancelled")],
        default="draft",
        tracking=True,
    )
    line_ids = fields.One2many("incentive.settlement.line", "settlement_id", copy=False)
    line_count = fields.Integer(compute="_compute_totals")
    total_base_amount = fields.Monetary(currency_field="company_currency_id", compute="_compute_totals", store=True)
    total_bonus_amount = fields.Monetary(currency_field="company_currency_id", compute="_compute_totals", store=True)
    total_special_reward_amount = fields.Monetary(currency_field="company_currency_id", compute="_compute_totals", store=True)
    total_final_amount = fields.Monetary(currency_field="company_currency_id", compute="_compute_totals", store=True)
    company_currency_id = fields.Many2one(related="company_id.currency_id", store=True, readonly=True)
    note = fields.Text()

    _sql_constraints = [
        ("incentive_settlement_period_uniq", "unique(company_id, period_id)", "A settlement already exists for this period and company."),
    ]

    @api.model_create_multi
    def create(self, vals_list):
        seq = self.env["ir.sequence"]
        for vals in vals_list:
            if vals.get("name", "New") == "New":
                vals["name"] = seq.next_by_code("incentive.settlement") or "New"
        return super().create(vals_list)

    @api.depends("line_ids.final_payable_amount_pkr", "line_ids.base_amount_pkr", "line_ids.bonus_amount_pkr", "line_ids.special_reward_amount_pkr")
    def _compute_totals(self):
        for rec in self:
            rec.line_count = len(rec.line_ids)
            rec.total_base_amount = sum(rec.line_ids.mapped("base_amount_pkr"))
            rec.total_bonus_amount = sum(rec.line_ids.mapped("bonus_amount_pkr"))
            rec.total_special_reward_amount = sum(rec.line_ids.mapped("special_reward_amount_pkr"))
            rec.total_final_amount = sum(rec.line_ids.mapped("final_payable_amount_pkr"))

    def _get_overlapping_target_domain(self):
        self.ensure_one()
        return [
            ("company_id", "=", self.company_id.id),
            ("state", "=", "confirmed"),
            ("review_period_id.date_start", "<=", self.date_end),
            ("review_period_id.date_end", ">=", self.date_start),
        ]

    def _get_source_domain(self):
        self.ensure_one()
        return [
            ("company_id", "=", self.company_id.id),
            ("state", "=", "confirmed"),
            ("receipt_date", ">=", self.date_start),
            ("receipt_date", "<=", self.date_end),
        ]

    def _pick_policy(self, employee, segment):
        return self.env["incentive.policy"].search(
            [
                ("company_id", "=", self.company_id.id),
                ("category_id", "=", employee.incentive_category_id.id),
                ("segment_id", "=", segment.id),
                ("active", "=", True),
            ],
            limit=1,
        )

    def _pick_slab(self, policy, achievement_percent):
        if not policy:
            return self.env["incentive.policy.line"]
        return policy.line_ids.sorted(key=lambda x: (x.achievement_from, x.id)).filtered(lambda line: line.matches(achievement_percent))[:1]

    def _prepare_employee_segment_metrics(self):
        self.ensure_one()
        Target = self.env["incentive.target.line"]
        Source = self.env["incentive.source.record"]
        metrics = defaultdict(lambda: {
            "target_total": 0.0,
            "achievement_total": 0.0,
            "target_lines": self.env["incentive.target.line"],
            "sources": self.env["incentive.source.record"],
        })

        for target in Target.search(self._get_overlapping_target_domain()):
            key = (target.employee_id.id, target.segment_id.id)
            metrics[key]["target_total"] += target.target_amount or 0.0
            metrics[key]["achievement_total"] += target.achievement_amount or 0.0
            metrics[key]["target_lines"] |= target

        for src in Source.search(self._get_source_domain()):
            key = (src.employee_id.id, src.segment_id.id)
            metrics[key]["sources"] |= src

        return metrics

    def action_generate_lines(self):
        Line = self.env["incentive.settlement.line"]
        for settlement in self:
            if settlement.state != "draft":
                raise UserError("Settlement lines can only be generated in draft state.")
            settlement.line_ids.unlink()
            metrics = settlement._prepare_employee_segment_metrics()
            created_lines = self.env["incentive.settlement.line"]

            for (employee_id, segment_id), data in metrics.items():
                employee = self.env["hr.employee"].browse(employee_id)
                segment = self.env["incentive.segment"].browse(segment_id)
                if not employee.incentive_category_id:
                    continue

                policy = settlement._pick_policy(employee, segment)
                target_total = data["target_total"]
                achievement_total = data["achievement_total"]
                achievement_percent = (achievement_total / target_total * 100.0) if target_total else 0.0
                slab = settlement._pick_slab(policy, achievement_percent)

                vals = settlement._prepare_segment_line_vals(
                    employee=employee,
                    segment=segment,
                    policy=policy,
                    slab=slab,
                    target_total=target_total,
                    achievement_total=achievement_total,
                    sources=data["sources"],
                )
                if vals:
                    created_lines |= Line.create(vals)

            settlement._generate_all_segments_bonus_lines(created_lines)
        return True

    def _prepare_segment_line_vals(self, employee, segment, policy, slab, target_total, achievement_total, sources):
        self.ensure_one()
        achievement_percent = (achievement_total / target_total * 100.0) if target_total else 0.0
        gross_salary = employee._get_incentive_gross_salary()
        category = employee.incentive_category_id

        eligibility_status = "eligible"
        block_reason = False
        source_count = len(sources)
        exception_status = "na"
        base_source_currency = 0.0
        base_amount_pkr = 0.0
        special_reward_amount_pkr = 0.0

        if not policy:
            eligibility_status = "blocked"
            block_reason = "No active incentive policy found."
        elif not slab and segment.basis_type != "commission_flat":
            eligibility_status = "blocked"
            block_reason = "No achievement slab matched the achievement percentage."

        if eligibility_status == "eligible":
            if segment.basis_type in ("commission_net", "profit_excl_gst", "commission_flat"):
                if not sources:
                    eligibility_status = "blocked"
                    block_reason = "No confirmed source records found in the settlement period."
                else:
                    for src in sources:
                        if not src.eligible:
                            continue
                        src_base = src.assessed_basis_amount
                        line_amount = 0.0
                        if slab:
                            if slab.calculation_type == "percent":
                                line_amount = src_base * ((slab.percentage or 0.0) / 100.0)
                            elif slab.calculation_type == "fixed_amount":
                                line_amount = slab.fixed_amount or 0.0
                        base_source_currency += line_amount
                        rate = src.exchange_rate or 1.0
                        base_amount_pkr += line_amount if src.currency_id == src.company_currency_id else line_amount * rate
                        if policy.special_reward_enabled and src.special_reward_eligible:
                            reward = src_base * ((policy.special_reward_percent or 0.0) / 100.0)
                            special_reward_amount_pkr += reward if src.currency_id == src.company_currency_id else reward * rate
                    if not base_amount_pkr and source_count:
                        exception_status = "pending_or_blocked"
            elif segment.basis_type == "gross_salary":
                if not slab:
                    eligibility_status = "blocked"
                    block_reason = "No salary slab matched the achievement percentage."
                elif slab.calculation_type == "salary_multiplier":
                    base_amount_pkr = gross_salary * (slab.salary_multiplier or 0.0)
                elif slab.calculation_type == "fixed_amount":
                    base_amount_pkr = slab.fixed_amount or 0.0
                else:
                    base_amount_pkr = gross_salary * ((slab.percentage or 0.0) / 100.0)

        final_amount = base_amount_pkr + special_reward_amount_pkr
        return {
            "settlement_id": self.id,
            "line_type": "segment",
            "employee_id": employee.id,
            "category_id": category.id,
            "segment_id": segment.id,
            "policy_id": policy.id if policy else False,
            "policy_line_id": slab.id if slab else False,
            "target_total": target_total,
            "achievement_total": achievement_total,
            "achievement_percent": achievement_percent,
            "source_record_ids": [(6, 0, sources.ids)] if sources else False,
            "source_record_count": source_count,
            "gross_salary_amount": gross_salary,
            "base_amount_source_currency": base_source_currency,
            "base_amount_pkr": base_amount_pkr,
            "bonus_amount_pkr": 0.0,
            "special_reward_amount_pkr": special_reward_amount_pkr,
            "final_payable_amount_pkr": final_amount,
            "eligibility_status": eligibility_status,
            "exception_status": exception_status,
            "blocked_reason": block_reason,
            "calculation_breakdown": self._build_calculation_breakdown(employee, segment, policy, slab, achievement_percent, base_source_currency, base_amount_pkr, special_reward_amount_pkr, gross_salary, sources),
        }

    def _build_calculation_breakdown(self, employee, segment, policy, slab, achievement_percent, base_amount_source_currency, base_amount_pkr, special_reward_amount_pkr, gross_salary, sources):
        return (
            f"Employee: {employee.name}\n"
            f"Segment: {segment.name}\n"
            f"Policy: {policy.name if policy else 'None'}\n"
            f"Slab: {slab.name if slab else 'None'}\n"
            f"Achievement %: {achievement_percent:.2f}\n"
            f"Gross Salary Basis: {gross_salary:.2f}\n"
            f"Base Amount (source currency): {base_amount_source_currency:.2f}\n"
            f"Base Amount (company currency): {base_amount_pkr:.2f}\n"
            f"Special Reward (company currency): {special_reward_amount_pkr:.2f}\n"
            f"Source References: {', '.join(sources.mapped('reference')) if sources else ''}"
        )

    def _generate_all_segments_bonus_lines(self, created_lines):
        self.ensure_one()
        employee_lines = defaultdict(lambda: self.env["incentive.settlement.line"])
        for line in created_lines.filtered(lambda l: l.line_type == "segment"):
            employee_lines[line.employee_id.id] |= line

        Line = self.env["incentive.settlement.line"]
        for employee_id, lines in employee_lines.items():
            employee = self.env["hr.employee"].browse(employee_id)
            eligible_lines = lines.filtered(lambda l: l.eligibility_status == "eligible" and l.segment_id and l.policy_id and l.policy_id.bonus_if_all_segments_100)
            if eligible_lines and all((line.achievement_percent or 0.0) >= 100.0 for line in eligible_lines):
                salary = employee._get_incentive_gross_salary()
                if salary:
                    Line.create({
                        "settlement_id": self.id,
                        "line_type": "bonus",
                        "employee_id": employee.id,
                        "category_id": employee.incentive_category_id.id,
                        "gross_salary_amount": salary,
                        "bonus_amount_pkr": salary,
                        "final_payable_amount_pkr": salary,
                        "eligibility_status": "eligible",
                        "exception_status": "na",
                        "calculation_breakdown": "Extra gross salary bonus applied because all eligible segments achieved at least 100%.",
                    })

    def action_review(self):
        self.write({"state": "reviewed"})

    def action_approve(self):
        self.write({"state": "approved"})

    def action_mark_paid(self):
        self.write({"state": "paid"})

    def action_set_draft(self):
        self.write({"state": "draft"})

    def action_cancel(self):
        self.write({"state": "cancelled"})


class IncentiveSettlementLine(models.Model):
    _name = "incentive.settlement.line"
    _description = "Incentive Settlement Line"
    _inherit = ["mail.thread", "mail.activity.mixin"]
    _order = "employee_id, line_type, segment_id"

    settlement_id = fields.Many2one("incentive.settlement", required=True, ondelete="cascade")
    line_type = fields.Selection([("segment", "Segment"), ("bonus", "Bonus")], default="segment", required=True)
    employee_id = fields.Many2one("hr.employee", required=True, tracking=True)
    category_id = fields.Many2one("incentive.employee.category", tracking=True)
    segment_id = fields.Many2one("incentive.segment", tracking=True)
    policy_id = fields.Many2one("incentive.policy", tracking=True)
    policy_line_id = fields.Many2one("incentive.policy.line", tracking=True)
    source_record_ids = fields.Many2many("incentive.source.record", "incentive_settlement_line_source_rel", "line_id", "source_id", string="Source Records")
    source_record_count = fields.Integer()

    target_total = fields.Float()
    achievement_total = fields.Float()
    achievement_percent = fields.Float()
    gross_salary_amount = fields.Monetary(currency_field="company_currency_id")
    base_amount_source_currency = fields.Float()
    base_amount_pkr = fields.Monetary(currency_field="company_currency_id")
    bonus_amount_pkr = fields.Monetary(currency_field="company_currency_id")
    special_reward_amount_pkr = fields.Monetary(currency_field="company_currency_id")
    manual_override_active = fields.Boolean(tracking=True)
    manual_override_amount_pkr = fields.Monetary(currency_field="company_currency_id", tracking=True)
    manual_override_reason = fields.Text(tracking=True)
    final_payable_amount_pkr = fields.Monetary(currency_field="company_currency_id", tracking=True)

    company_currency_id = fields.Many2one(related="settlement_id.company_currency_id", store=True, readonly=True)
    eligibility_status = fields.Selection([("eligible", "Eligible"), ("blocked", "Blocked")], default="eligible", tracking=True)
    exception_status = fields.Selection([("na", "Not Applicable"), ("pending_or_blocked", "Pending / Blocked"), ("approved", "Approved")], default="na", tracking=True)
    blocked_reason = fields.Char(tracking=True)
    calculation_breakdown = fields.Text()
    company_id = fields.Many2one(related="settlement_id.company_id", store=True, readonly=True)

    @api.constrains("manual_override_active", "manual_override_reason")
    def _check_manual_override_reason(self):
        for rec in self:
            if rec.manual_override_active and not rec.manual_override_reason:
                raise ValidationError("Manual override reason is required.")

    def _recompute_final_amount(self):
        for rec in self:
            if rec.manual_override_active:
                rec.final_payable_amount_pkr = rec.manual_override_amount_pkr
            else:
                rec.final_payable_amount_pkr = (rec.base_amount_pkr or 0.0) + (rec.bonus_amount_pkr or 0.0) + (rec.special_reward_amount_pkr or 0.0)

    @api.onchange("manual_override_active", "manual_override_amount_pkr")
    def _onchange_manual_override(self):
        self._recompute_final_amount()

    def write(self, vals):
        res = super().write(vals)
        tracked = {"manual_override_active", "manual_override_amount_pkr", "base_amount_pkr", "bonus_amount_pkr", "special_reward_amount_pkr"}
        if tracked & set(vals.keys()):
            self._recompute_final_amount()
        return res

    def action_open_sources(self):
        self.ensure_one()
        return {
            "type": "ir.actions.act_window",
            "name": "Source Records",
            "res_model": "incentive.source.record",
            "view_mode": "list,form",
            "domain": [("id", "in", self.source_record_ids.ids)],
        }
