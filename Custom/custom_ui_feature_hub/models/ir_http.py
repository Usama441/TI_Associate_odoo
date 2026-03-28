from odoo import models


def _is_true(value):
    return str(value).strip().lower() in {'1', 'true', 'yes', 'on'}


class IrHttp(models.AbstractModel):
    _inherit = 'ir.http'

    def session_info(self):
        result = super().session_info()
        if not self.env.user._is_internal():
            return result

        config = self.env['ir.config_parameter'].sudo()
        result['ui_feature_hub'] = {
            'theme_preset': config.get_param('custom_ui_feature_hub.theme_preset', default='default'),
            'announcement_enabled': _is_true(config.get_param('custom_ui_feature_hub.announcement_enabled', default='False')),
            'announcement_message': config.get_param('custom_ui_feature_hub.announcement_message', default=''),
            'quick_links_enabled': _is_true(config.get_param('custom_ui_feature_hub.quick_links_enabled', default='False')),
            'quick_links_text': config.get_param('custom_ui_feature_hub.quick_links_text', default=''),
            'show_settings_summary': _is_true(config.get_param('custom_ui_feature_hub.show_settings_summary', default='True')),
            'show_effect': _is_true(config.get_param('base_setup.show_effect', default='False')),
            'company_name': self.env.company.display_name,
            'user_name': self.env.user.name,
        }
        return result
