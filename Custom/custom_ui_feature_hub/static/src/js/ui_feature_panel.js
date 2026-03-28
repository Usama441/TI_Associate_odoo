import { registry } from "@web/core/registry";
import { session } from "@web/session";
import { Component } from "@odoo/owl";

function parseQuickLinks(textValue) {
    const value = (textValue || "").trim();
    if (!value) {
        return [];
    }
    return value
        .split(/[\n;]+/)
        .map((line) => line.trim())
        .filter(Boolean)
        .map((line) => {
            const [label, url] = line.split("|");
            return {
                label: (label || "").trim(),
                url: (url || "").trim(),
            };
        })
        .filter((item) => item.label && item.url);
}

export class UIFeaturePanel extends Component {
    static template = "custom_ui_feature_hub.UIFeaturePanel";
    static props = {};

    setup() {
        this.uiConfig = session.ui_feature_hub || {};
    }

    get showSummary() {
        return Boolean(this.uiConfig.show_settings_summary);
    }

    get quickLinks() {
        if (!this.uiConfig.quick_links_enabled) {
            return [];
        }
        return parseQuickLinks(this.uiConfig.quick_links_text);
    }

    get isVisible() {
        return this.showSummary || this.quickLinks.length;
    }

    get summaryItems() {
        return [
            { label: "Company", value: this.uiConfig.company_name || "-" },
            { label: "User", value: this.uiConfig.user_name || "-" },
            { label: "Theme", value: this.uiConfig.theme_preset || "default" },
        ];
    }
}

registry.category("main_components").add(
    "custom_ui_feature_hub.UIFeaturePanel",
    {
        Component: UIFeaturePanel,
    },
    { sequence: 80 }
);
