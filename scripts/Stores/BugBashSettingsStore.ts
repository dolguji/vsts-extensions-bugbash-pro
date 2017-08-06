import { BaseStore } from "VSTS_Extension/Flux/Stores/BaseStore";

import { SettingsActionsHub } from "../Actions/ActionsHub";
import { BugBashSettings } from "../Interfaces";

export class BugBashSettingsStore extends BaseStore<BugBashSettings, BugBashSettings, void> {
    public getItem(_id: void): BugBashSettings {
         return this.items;
    }

    protected initializeActionListeners() {
        SettingsActionsHub.InitializeBugBashSettings.addListener((settings: BugBashSettings) => {
            if (settings) {
                this.items = settings;
            }

            this.emitChanged();
        });

        SettingsActionsHub.UpdateBugBashSettings.addListener((settings: BugBashSettings) => {
            if (settings) {
                this.items = settings;
            }

            this.emitChanged();
        });
    }

    public getKey(): string {
        return "BugBashSettingsStore";
    }

    protected convertItemKeyToString(_key: void): string {
        return null;
    }
}