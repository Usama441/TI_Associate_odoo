{
    'name': 'Custom UI Feature Hub',
    'version': '19.0.1.0.0',
    'summary': 'Backend UI presets and user-visible widgets',
    'description': """
        Adds lightweight backend UI presets, announcement banner,
        quick links, and read-only settings summary widgets for internal users.
    """,
    'author': 'TI Associates',
    'license': 'LGPL-3',
    'category': 'Tools',
    'depends': [
        'base',
        'web',
        'base_setup',
    ],
    'data': [
        'views/res_config_settings_views.xml',
    ],
    'assets': {
        'web.assets_backend': [
            'custom_ui_feature_hub/static/src/scss/ui_feature_backend.scss',
            'custom_ui_feature_hub/static/src/js/ui_feature_banner.js',
            'custom_ui_feature_hub/static/src/js/ui_feature_panel.js',
            'custom_ui_feature_hub/static/src/js/ui_feature_systray.js',
            'custom_ui_feature_hub/static/src/xml/ui_feature_widgets.xml',
        ],
    },
    'installable': True,
    'application': False,
}
