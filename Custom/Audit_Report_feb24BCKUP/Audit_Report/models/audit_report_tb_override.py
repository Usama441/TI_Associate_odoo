from odoo import api, fields, models


TB_PERIOD_SELECTION = [
    ('current', 'Current'),
    ('prior', 'Prior'),
]


class AuditReportTbOverrideLine(models.TransientModel):
    _name = 'audit.report.tb.override.line'
    _description = 'Audit Report Trial Balance Override Line'
    _order = 'period_key, account_code, id'

    wizard_id = fields.Many2one(
        'audit.report',
        required=True,
        ondelete='cascade',
        index=True,
    )
    period_key = fields.Selection(
        TB_PERIOD_SELECTION,
        required=True,
        index=True,
    )
    account_id = fields.Many2one(
        'account.account',
        string='Account',
        readonly=True,
        index=True,
    )
    account_code = fields.Char(
        string='Code',
        required=True,
        readonly=True,
    )
    account_name = fields.Char(
        string='Name',
        readonly=True,
    )

    system_debit = fields.Float(string='System Debit', digits='Account', readonly=True)
    system_credit = fields.Float(string='System Credit', digits='Account', readonly=True)
    system_balance = fields.Float(string='System Balance', digits='Account', readonly=True)

    override_debit = fields.Float(string='Override Debit', digits='Account')
    override_credit = fields.Float(string='Override Credit', digits='Account')
    override_balance = fields.Float(string='Override Balance', digits='Account')

    effective_debit = fields.Float(
        string='Effective Debit',
        digits='Account',
        compute='_compute_effective_amounts',
        store=True,
        readonly=True,
    )
    effective_credit = fields.Float(
        string='Effective Credit',
        digits='Account',
        compute='_compute_effective_amounts',
        store=True,
        readonly=True,
    )
    effective_balance = fields.Float(
        string='Effective Balance',
        digits='Account',
        compute='_compute_effective_amounts',
        store=True,
        readonly=True,
    )
    is_overridden = fields.Boolean(
        string='Overridden',
        compute='_compute_is_overridden',
        store=True,
        readonly=True,
    )

    @staticmethod
    def _to_float(value):
        try:
            return float(value or 0.0)
        except (TypeError, ValueError):
            return 0.0

    @staticmethod
    def _is_different(left, right):
        return abs((left or 0.0) - (right or 0.0)) > 1e-6

    @api.model
    def _normalize_override_vals(self, vals, record=None):
        normalized = dict(vals or {})

        default_debit = self._to_float(
            normalized.get('system_debit', record.system_debit if record else 0.0)
        )
        default_credit = self._to_float(
            normalized.get('system_credit', record.system_credit if record else 0.0)
        )

        if 'override_debit' not in normalized:
            normalized['override_debit'] = self._to_float(
                record.override_debit if record else default_debit
            )
        if 'override_credit' not in normalized:
            normalized['override_credit'] = self._to_float(
                record.override_credit if record else default_credit
            )

        if 'override_balance' in normalized and 'override_credit' not in vals:
            debit_value = self._to_float(normalized.get('override_debit'))
            balance_value = self._to_float(normalized.get('override_balance'))
            normalized['override_credit'] = debit_value - balance_value

        debit_value = self._to_float(normalized.get('override_debit'))
        credit_value = self._to_float(normalized.get('override_credit'))
        normalized['override_balance'] = debit_value - credit_value

        return normalized

    @api.model_create_multi
    def create(self, vals_list):
        normalized_vals_list = [self._normalize_override_vals(vals) for vals in vals_list]
        return super().create(normalized_vals_list)

    def write(self, vals):
        if not any(
            key in vals
            for key in ('override_debit', 'override_credit', 'override_balance', 'system_debit', 'system_credit')
        ):
            return super().write(vals)

        if len(self) == 1:
            normalized_vals = self._normalize_override_vals(vals, record=self)
            return super().write(normalized_vals)

        for record in self:
            normalized_vals = self._normalize_override_vals(vals, record=record)
            super(AuditReportTbOverrideLine, record).write(normalized_vals)
        return True

    @api.onchange('override_debit', 'override_credit')
    def _onchange_override_debit_credit(self):
        for record in self:
            record.override_balance = (
                self._to_float(record.override_debit)
                - self._to_float(record.override_credit)
            )

    @api.onchange('override_balance')
    def _onchange_override_balance(self):
        for record in self:
            debit_value = self._to_float(record.override_debit)
            balance_value = self._to_float(record.override_balance)
            # Balance edit keeps debit as source and auto-adjusts credit.
            record.override_credit = debit_value - balance_value
            record.override_balance = debit_value - self._to_float(record.override_credit)

    @api.depends('override_debit', 'override_credit')
    def _compute_effective_amounts(self):
        for record in self:
            debit_value = self._to_float(record.override_debit)
            credit_value = self._to_float(record.override_credit)
            record.effective_debit = debit_value
            record.effective_credit = credit_value
            record.effective_balance = debit_value - credit_value

    @api.depends(
        'effective_debit',
        'effective_credit',
        'effective_balance',
        'system_debit',
        'system_credit',
        'system_balance',
    )
    def _compute_is_overridden(self):
        for record in self:
            record.is_overridden = any([
                self._is_different(record.effective_debit, record.system_debit),
                self._is_different(record.effective_credit, record.system_credit),
                self._is_different(record.effective_balance, record.system_balance),
            ])
