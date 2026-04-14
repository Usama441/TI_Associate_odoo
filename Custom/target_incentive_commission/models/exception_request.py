from odoo import api, fields, models
from odoo.exceptions import ValidationError


class IncentiveExceptionRequest(models.Model):
    _name = "incentive.exception.request"
    _description = "Incentive Exception / Approval Request"
    _inherit = ["mail.thread", "mail.activity.mixin"]
    _order = "create_date desc, id desc"

    name = fields.Char(required=True, copy=False, default=lambda self: self.env["ir.sequence"].next_by_code("incentive.exception.request") or "New")
    company_id = fields.Many2one("res.company", default=lambda self: self.env.company, required=True)
    source_record_id = fields.Many2one("incentive.source.record", required=True, ondelete="cascade", tracking=True)
    employee_id = fields.Many2one(related="source_record_id.employee_id", store=True, readonly=True)
    segment_id = fields.Many2one(related="source_record_id.segment_id", store=True, readonly=True)
    exception_type = fields.Selection(
        [("gp_below_threshold", "GP Below Threshold"), ("manual_override", "Manual Override"), ("special_case", "Special Case")],
        required=True,
        default="gp_below_threshold",
        tracking=True,
    )
    requested_by = fields.Many2one("res.users", default=lambda self: self.env.user, required=True, tracking=True)
    reason = fields.Text(required=True, tracking=True)
    requested_override_amount = fields.Monetary(currency_field="company_currency_id")
    company_currency_id = fields.Many2one(related="company_id.currency_id", store=True, readonly=True)
    state = fields.Selection(
        [("draft", "Draft"), ("submitted", "Submitted"), ("approved", "Approved"), ("rejected", "Rejected"), ("cancelled", "Cancelled")],
        default="draft",
        tracking=True,
    )
    approved_by = fields.Many2one("res.users", tracking=True)
    approved_on = fields.Datetime(tracking=True)
    manager_note = fields.Text()

    @api.constrains("source_record_id")
    def _check_duplicate_open_request(self):
        for rec in self:
            domain = [
                ("id", "!=", rec.id),
                ("source_record_id", "=", rec.source_record_id.id),
                ("exception_type", "=", rec.exception_type),
                ("state", "in", ["draft", "submitted", "approved"]),
            ]
            if self.search_count(domain):
                raise ValidationError("Another active request already exists for this source record and exception type.")

    def action_submit(self):
        self.write({"state": "submitted"})

    def action_approve(self):
        self.write({"state": "approved", "approved_by": self.env.user.id, "approved_on": fields.Datetime.now()})

    def action_reject(self):
        self.write({"state": "rejected"})

    def action_cancel(self):
        self.write({"state": "cancelled"})
