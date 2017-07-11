import { ExtensionDataManager } from "VSTS_Extension/Utilities/ExtensionDataManager";

import { SettingsActionsCreator } from "./ActionsCreator";
import { SettingsStore } from "../Stores/SettingsStore";
import { Settings } from "../Interfaces";
import { StoresHub } from "../Stores/StoresHub";

export module SettingsActions {
    export async function initializeBugBashSettings() {
        if (StoresHub.settingsStore.isLoaded()) {
            SettingsActionsCreator.InitializeBugBashSettings.invoke(null);
        }
        else if (!StoresHub.settingsStore.isLoading()) {
            StoresHub.settingsStore.setLoading(true);

            try {
                const settings = await ExtensionDataManager.readUserSetting(`bugBashProSettings_${VSS.getWebContext().project.id}`, {} as Settings, false);
                
                SettingsActionsCreator.InitializeBugBashSettings.invoke(settings);
                StoresHub.settingsStore.setLoading(false);
            }
            catch (e) {
                StoresHub.settingsStore.setLoading(false);
                throw e;
            }
        }
    }

    export async function updateBugBashSettings(settings: Settings) {        
        StoresHub.settingsStore.setLoading(true);
        
        try {
            const updatedSettings = await ExtensionDataManager.writeUserSetting<Settings>(`bugBashProSettings_${VSS.getWebContext().project.id}`, settings, false);
            
            SettingsActionsCreator.UpdateBugBashSettings.invoke(updatedSettings);
            StoresHub.settingsStore.setLoading(false);
        }
        catch (e) {
            StoresHub.settingsStore.setLoading(false);
            throw e;
        }
    }
}