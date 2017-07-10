import { BaseStore } from "VSTS_Extension/Flux/Stores/BaseStore";

import { SettingsActionsCreator } from "../Actions/ActionsCreator";
import { Settings } from "../Interfaces";

export class SettingsStore extends BaseStore<Settings, Settings, void> {
    public getItem(id: void): Settings {
         return this.items;
    }

    protected initializeActionListeners() {
        SettingsActionsCreator.InitializeBugBashSettings.addListener((settings: Settings) => {
            if (settings) {
                this.items = settings;
            }

            this.emitChanged();
        });

        SettingsActionsCreator.UpdateBugBashSettings.addListener((settings: Settings) => {
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