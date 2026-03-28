import { Dropdown } from "@web/core/dropdown/dropdown";
import { DropdownGroup } from "@web/core/dropdown/dropdown_group";
import { DropdownItem } from "@web/core/dropdown/dropdown_item";
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

export class UIFeatureSystray extends Component {
    static template = "custom_ui_feature_hub.UIFeatureSystray";
    static components = { Dropdown, DropdownGroup, DropdownItem };
    static props = {};

    setup() {
        this.uiConfig = session.ui_feature_hub || {};
    }

    get quickLinks() {
        return parseQuickLinks(this.uiConfig.quick_links_text);
    }

    get hasQuickLinks() {
        return Boolean(this.uiConfig.quick_links_enabled && this.quickLinks.length);
    }

    get showSummary() {
        return Boolean(this.uiConfig.show_settings_summary);
    }

    get summaryItems() {
        return [
            {
                label: "Company",
                value: this.uiConfig.company_name || "-",
            },
            {
                label: "Style Preset",
                value: this.uiConfig.theme_preset || "default",
            },
            {
                label: "Announcement",
                value: this.uiConfig.announcement_enabled ? "Enabled" : "Disabled",
            },
            {
                label: "Show Effects",
                value: this.uiConfig.show_effect ? "Yes" : "No",
            },
        ];
    }

    get isVisible() {
        return this.hasQuickLinks || this.showSummary;
    }
}

registry.category("systray").add("custom_ui_feature_hub.UIFeatureSystray", {
    Component: UIFeatureSystray,
}, { sequence: 35 });
