import { ExtensionDataManager } from "VSTS_Extension/Utilities/ExtensionDataManager";

import { SettingsActionsHub } from "./ActionsHub";
import { IUserSettings, IBugBashSettings } from "../Interfaces";
import { StoresHub } from "../Stores/StoresHub";

export module SettingsActions {
    export async function initializeBugBashSettings() {
        if (StoresHub.bugBashSettingsStore.isLoaded()) {
            SettingsActionsHub.InitializeBugBashSettings.invoke(null);
        }
        else if (!StoresHub.bugBashSettingsStore.isLoading()) {
            StoresHub.bugBashSettingsStore.setLoading(true);
            const settings = await ExtensionDataManager.readUserSetting(`bugBashProSettings_${VSS.getWebContext().project.id}`, {} as IBugBashSettings, false);            
            SettingsActionsHub.InitializeBugBashSettings.invoke(settings);
            StoresHub.bugBashSettingsStore.setLoading(false);
        }
    }

    export async function updateBugBashSettings(settings: IBugBashSettings) {
        try {
            const updatedSettings = await ExtensionDataManager.writeUserSetting<IBugBashSettings>(`bugBashProSettings_${VSS.getWebContext().project.id}`, settings, false);
            SettingsActionsHub.UpdateBugBashSettings.invoke(updatedSettings);
        }
        catch (e) {
            throw e.message;
        }
    }

    export async function initializeUserSettings() {
        if (StoresHub.userSettingsStore.isLoaded()) {
            SettingsActionsHub.InitializeUserSettings.invoke(null);
        }
        else if (!StoresHub.userSettingsStore.isLoading()) {
            StoresHub.userSettingsStore.setLoading(true);
            const settings = await ExtensionDataManager.readDocuments<IUserSettings>(`UserSettings_${VSS.getWebContext().project.id}`, false);
            SettingsActionsHub.InitializeUserSettings.invoke(settings);
            StoresHub.userSettingsStore.setLoading(false);
        }
    }

    export async function updateUserSettings(settings: IUserSettings) {
        try {
            const updatedSettings = await ExtensionDataManager.addOrUpdateDocument<IUserSettings>(`UserSettings_${VSS.getWebContext().project.id}`, settings, false);
            SettingsActionsHub.UpdateUserSettings.invoke(updatedSettings);
        }
        catch (e) {
            throw e.message;
        }
    }
}