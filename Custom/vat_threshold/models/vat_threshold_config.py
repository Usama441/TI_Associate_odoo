# -*- coding: utf-8 -*-

from odoo import models, fields, api


class VatThresholdConfig(models.TransientModel):
    _name = 'vat.threshold.config'
    _description = 'VAT Threshold Settings'
    _inherit = 'res.config.settings'

    daily_report_recipients = fields.Char(
        string='Daily Report Email Recipients',
        config_parameter='vat_threshold.daily_report_recipients',
        help='Comma-separated email addresses for daily VAT threshold reports'
    )