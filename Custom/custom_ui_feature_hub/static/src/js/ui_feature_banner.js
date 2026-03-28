import { registry } from "@web/core/registry";
import { session } from "@web/session";
import { Component } from "@odoo/owl";

export class UIFeatureBanner extends Component {
    static template = "custom_ui_feature_hub.UIFeatureBanner";
    static props = {};

    setup() {
        this.uiConfig = session.ui_feature_hub || {};
        this._applyBodyPreset();
    }

    _applyBodyPreset() {
        const body = document.body;
        if (!body) {
            return;
        }
        for (const className of [...body.classList]) {
            if (className.startsWith("o_ui_feature_theme_")) {
                body.classList.remove(className);
            }
        }
        const preset = this.uiConfig.theme_preset || "default";
        body.classList.add(`o_ui_feature_theme_${preset}`);
    }

    get showBanner() {
        return Boolean(this.uiConfig.announcement_enabled && this.uiConfig.announcement_message);
    }

    get message() {
        return this.uiConfig.announcement_message || "";
    }
}

registry.category("main_components").add(
    "custom_ui_feature_hub.UIFeatureBanner",
    {
        Component: UIFeatureBanner,
    },
    { sequence: 10 }
);
