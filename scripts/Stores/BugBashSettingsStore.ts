import { BaseStore } from "MB/Flux/Stores/BaseStore";

import { SettingsActionsHub } from "../Actions/ActionsHub";
import { IBugBashSettings } from "../Interfaces";

export class BugBashSettingsStore extends BaseStore<IBugBashSettings, IBugBashSettings, void> {
    public getItem(_id: void): IBugBashSettings {
         return this.items;
    }

    protected initializeActionListeners() {
        SettingsActionsHub.InitializeBugBashSettings.addListener((settings: IBugBashSettings) => {
            if (settings) {
                this.items = settings;
            }

            this.emitChanged();
        });

        SettingsActionsHub.UpdateBugBashSettings.addListener((settings: IBugBashSettings) => {
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