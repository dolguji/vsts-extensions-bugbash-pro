import { ExtensionDataManager } from "VSTS_Extension/Utilities/ExtensionDataManager";

import { SettingsActionsHub } from "./ActionsHub";
import { SettingsStore } from "../Stores/SettingsStore";
import { Settings } from "../Interfaces";
import { StoresHub } from "../Stores/StoresHub";

export module SettingsActions {
    export async function initializeBugBashSettings() {
        if (StoresHub.settingsStore.isLoaded()) {
            SettingsActionsHub.InitializeBugBashSettings.invoke(null);
        }
        else if (!StoresHub.settingsStore.isLoading()) {
            StoresHub.settingsStore.setLoading(true);
            const settings = await ExtensionDataManager.readUserSetting(`bugBashProSettings_${VSS.getWebContext().project.id}`, {} as Settings, false);            
            SettingsActionsHub.InitializeBugBashSettings.invoke(settings);
            StoresHub.settingsStore.setLoading(false);
        }
    }

    export async function updateBugBashSettings(settings: Settings) {        
        StoresHub.settingsStore.setLoading(true);
        
        try {
            const updatedSettings = await ExtensionDataManager.writeUserSetting<Settings>(`bugBashProSettings_${VSS.getWebContext().project.id}`, settings, false);
            
            SettingsActionsHub.UpdateBugBashSettings.invoke(updatedSettings);
            StoresHub.settingsStore.setLoading(false);
        }
        catch (e) {
            StoresHub.settingsStore.setLoading(false);
            throw e.message;
        }
    }
}