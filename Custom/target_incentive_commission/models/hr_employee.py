from odoo import fields, models


class HrEmployee(models.Model):
    _inherit = "hr.employee"

    incentive_category_id = fields.Many2one("incentive.employee.category", tracking=True)
    incentive_manager_id = fields.Many2one("hr.employee", string="Incentive Approver / Manager", tracking=True)
    incentive_segment_ids = fields.Many2many(
        "incentive.segment",
        "hr_employee_incentive_segment_rel",
        "employee_id",
        "segment_id",
        string="Relevant Target Groups",
    )
    gross_salary_amount = fields.Monetary(currency_field="company_currency_id", tracking=True)
    salary_source_method = fields.Selection(
        [
            ("employee_field", "Employee Incentive Profile"),
            ("contract_if_available", "Active Contract if Available, else Profile"),
        ],
        default="contract_if_available",
        required=True,
        tracking=True,
    )
    company_currency_id = fields.Many2one("res.currency", related="company_id.currency_id", store=True, readonly=True)
    incentive_note = fields.Text()
    target_line_count = fields.Integer(compute="_compute_incentive_counts")
    source_record_count = fields.Integer(compute="_compute_incentive_counts")
    settlement_line_count = fields.Integer(compute="_compute_incentive_counts")

    def _compute_incentive_counts(self):
        target_data = self.env["incentive.target.line"].read_group([("employee_id", "in", self.ids)], ["employee_id"], ["employee_id"])
        source_data = self.env["incentive.source.record"].read_group([("employee_id", "in", self.ids)], ["employee_id"], ["employee_id"])
        settlement_data = self.env["incentive.settlement.line"].read_group([("employee_id", "in", self.ids)], ["employee_id"], ["employee_id"])
        target_map = {d["employee_id"][0]: d["employee_id_count"] for d in target_data}
        source_map = {d["employee_id"][0]: d["employee_id_count"] for d in source_data}
        settlement_map = {d["employee_id"][0]: d["employee_id_count"] for d in settlement_data}
        for emp in self:
            emp.target_line_count = target_map.get(emp.id, 0)
            emp.source_record_count = source_map.get(emp.id, 0)
            emp.settlement_line_count = settlement_map.get(emp.id, 0)

    def action_view_incentive_targets(self):
        self.ensure_one()
        return {
            "type": "ir.actions.act_window",
            "name": "Target Lines",
            "res_model": "incentive.target.line",
            "view_mode": "list,form",
            "domain": [("employee_id", "=", self.id)],
            "context": {"default_employee_id": self.id},
        }

    def action_view_incentive_sources(self):
        self.ensure_one()
        return {
            "type": "ir.actions.act_window",
            "name": "Source Records",
            "res_model": "incentive.source.record",
            "view_mode": "list,form",
            "domain": [("employee_id", "=", self.id)],
            "context": {"default_employee_id": self.id},
        }

    def action_view_incentive_settlements(self):
        self.ensure_one()
        return {
            "type": "ir.actions.act_window",
            "name": "Settlement Lines",
            "res_model": "incentive.settlement.line",
            "view_mode": "list,form",
            "domain": [("employee_id", "=", self.id)],
        }

    def _get_incentive_gross_salary(self):
        self.ensure_one()
        if self.salary_source_method == "contract_if_available" and "hr.contract" in self.env:
            Contract = self.env["hr.contract"]
            if "wage" in Contract._fields:
                contract = Contract.search([("employee_id", "=", self.id), ("state", "=", "open")], limit=1, order="date_start desc, id desc")
                if contract:
                    return contract.wage
        return self.gross_salary_amount or 0.0
