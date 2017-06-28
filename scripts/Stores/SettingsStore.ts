import { BaseStore } from "VSTS_Extension/Stores/BaseStore";
import { ExtensionDataManager } from "VSTS_Extension/Utilities/ExtensionDataManager";

import { Settings } from "../Interfaces";

export class SettingsStore extends BaseStore<Settings, Settings, void> {
    protected getItemByKey(id: void): Settings {
         return null;
    }

    protected async initializeItems(): Promise<void> {
        this.items = await ExtensionDataManager.readUserSetting(`bugBashProSettings_${VSS.getWebContext().project.id}`, {} as Settings, false);
    }

    public getKey(): string {
        return "SettingsStore";
    }

    public async updateSettings(settings: Settings): Promise<void> {
        this.items = await ExtensionDataManager.writeUserSetting<Settings>(`bugBashProSettings_${VSS.getWebContext().project.id}`, settings, false);
        this.emitChanged();
    }
}