import { StoreFactory } from "VSTS_Extension/Flux/Stores/BaseStore";
import { ExtensionDataManager } from "VSTS_Extension/Utilities/ExtensionDataManager";

import { SettingsActionsCreator } from "./ActionsCreator";
import { SettingsStore } from "../Stores/SettingsStore";
import { Settings } from "../Interfaces";

export module SettingsActions {
    var settingsStore: SettingsStore = StoreFactory.getInstance<SettingsStore>(SettingsStore);

    export async function initializeBugBashSettings() {
        if (settingsStore.isLoaded()) {
            SettingsActionsCreator.InitializeBugBashSettings.invoke(null);
        }
        else if (!settingsStore.isLoading()) {
            settingsStore.setLoading(true);

            try {
                const settings = await ExtensionDataManager.readUserSetting(`bugBashProSettings_${VSS.getWebContext().project.id}`, {} as Settings, false);
                settingsStore.setLoading(false);
                settingsStore.setError(null);

                SettingsActionsCreator.InitializeBugBashSettings.invoke(settings);
            }
            catch (e) {
                settingsStore.setLoading(false);
                settingsStore.setError(e.message || e);
            }
        }
    }

    export async function updateBugBashSettings(settings: Settings) {        
        settingsStore.setLoading(true);
        
        try {
            const updatedSettings = await ExtensionDataManager.writeUserSetting<Settings>(`bugBashProSettings_${VSS.getWebContext().project.id}`, settings, false);
            settingsStore.setLoading(false);
            settingsStore.setError(null);

            SettingsActionsCreator.UpdateBugBashSettings.invoke(updatedSettings);
        }
        catch (e) {
            settingsStore.setLoading(false);
            settingsStore.setError(e.message || e);
        }
    }
}