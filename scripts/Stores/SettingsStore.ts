import { BaseStore } from "VSTS_Extension/Flux/Stores/BaseStore";

import { SettingsActionsHub } from "../Actions/ActionsHub";
import { Settings } from "../Interfaces";

export class SettingsStore extends BaseStore<Settings, Settings, void> {
    public getItem(id: void): Settings {
         return this.items;
    }

    protected initializeActionListeners() {
        SettingsActionsHub.InitializeBugBashSettings.addListener((settings: Settings) => {
            if (settings) {
                this.items = settings;
            }

            this.emitChanged();
        });

        SettingsActionsHub.UpdateBugBashSettings.addListener((settings: Settings) => {
            if (settings) {
                this.items = settings;
            }

            this.emitChanged();
        });
    }

    public getKey(): string {
        return "SettingsStore";
    }

    protected convertItemKeyToString(key: void): string {
        return null;
    }
}