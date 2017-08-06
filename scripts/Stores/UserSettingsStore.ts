import { BaseStore } from "VSTS_Extension/Flux/Stores/BaseStore";

import { SettingsActionsHub } from "../Actions/ActionsHub";
import { UserSettings } from "../Interfaces";

export class UserSettingsStore extends BaseStore<UserSettings, UserSettings, void> {
    public getItem(_id: void): UserSettings {
         return this.items;
    }

    protected initializeActionListeners() {
        SettingsActionsHub.InitializeUserSettings.addListener((settings: UserSettings) => {
            if (settings) {
                this.items = settings;
            }

            this.emitChanged();
        });

        SettingsActionsHub.UpdateUserSettings.addListener((settings: UserSettings) => {
            if (settings) {
                this.items = settings;
            }

            this.emitChanged();
        });
    }

    public getKey(): string {
        return "ProjectSettingsStore";
    }

    protected convertItemKeyToString(_key: void): string {
        return null;
    }
}