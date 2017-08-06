import { BaseStore } from "VSTS_Extension/Flux/Stores/BaseStore";

import { SettingsActionsHub } from "../Actions/ActionsHub";
import { IUserSettings } from "../Interfaces";

export class UserSettingsStore extends BaseStore<IDictionaryStringTo<IUserSettings>, IUserSettings, string> {
    constructor() {
        super();
        this.items = {};    
    }

    public getItem(id: string): IUserSettings {
         return this.items[id.toLowerCase()];
    }

    protected initializeActionListeners() {
        SettingsActionsHub.InitializeUserSettings.addListener((settings: IUserSettings[]) => {
            if (settings) {
                for (const setting of settings) {
                    this.items[setting.id.toLowerCase()] = setting;
                }
            }

            this.emitChanged();
        });

        SettingsActionsHub.UpdateUserSettings.addListener((settings: IUserSettings) => {
            if (settings) {
                this.items[settings.id.toLowerCase()] = settings;
            }

            this.emitChanged();
        });
    }

    public getKey(): string {
        return "UserSettingsStore";
    }

    protected convertItemKeyToString(key: string): string {
        return key;
    }
}