from odoo import fields, models


class ResConfigSettings(models.TransientModel):
    _inherit = 'res.config.settings'

    ui_feature_theme_preset = fields.Selection(
        selection=[
            ('default', 'Default'),
            ('compact', 'Compact'),
            ('contrast', 'High Contrast'),
            ('ramadan', 'Ramadan (Gold & Black)'),
            ('ramadan_soft', 'Ramadan Soft (Gold & Night)'),
        ],
        string='Backend Style Preset',
        default='default',
        config_parameter='custom_ui_feature_hub.theme_preset',
    )
    ui_feature_announcement_enabled = fields.Boolean(
        string='Enable Announcement Banner',
        config_parameter='custom_ui_feature_hub.announcement_enabled',
    )
    ui_feature_announcement_message = fields.Char(
        string='Announcement Message',
        config_parameter='custom_ui_feature_hub.announcement_message',
    )
    ui_feature_quick_links_enabled = fields.Boolean(
        string='Enable Quick Links Widget',
        config_parameter='custom_ui_feature_hub.quick_links_enabled',
    )
    ui_feature_quick_links_text = fields.Char(
        string='Quick Links (format: Label|URL ; Label|URL)',
        config_parameter='custom_ui_feature_hub.quick_links_text',
        default='Home|/odoo ; My Profile|/odoo/my ; Apps|/odoo/apps',
    )
    ui_feature_show_settings_summary = fields.Boolean(
        string='Show Settings Summary to Internal Users',
        config_parameter='custom_ui_feature_hub.show_settings_summary',
        default=True,
    )
