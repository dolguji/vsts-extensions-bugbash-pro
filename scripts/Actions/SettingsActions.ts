import { ExtensionDataManager } from "VSTS_Extension/Utilities/ExtensionDataManager";

import { SettingsActionsHub } from "./ActionsHub";
import { UserSettings, BugBashSettings } from "../Interfaces";
import { StoresHub } from "../Stores/StoresHub";

export module SettingsActions {
    export async function initializeBugBashSettings() {
        if (StoresHub.bugBashSettingsStore.isLoaded()) {
            SettingsActionsHub.InitializeBugBashSettings.invoke(null);
        }
        else if (!StoresHub.bugBashSettingsStore.isLoading()) {
            StoresHub.bugBashSettingsStore.setLoading(true);
            const settings = await ExtensionDataManager.readUserSetting(`bugBashProSettings_${VSS.getWebContext().project.id}`, {} as BugBashSettings, false);            
            SettingsActionsHub.InitializeBugBashSettings.invoke(settings);
            StoresHub.bugBashSettingsStore.setLoading(false);
        }
    }

    export async function updateBugBashSettings(settings: BugBashSettings) {        
        StoresHub.bugBashSettingsStore.setLoading(true);
        
        try {
            const updatedSettings = await ExtensionDataManager.writeUserSetting<BugBashSettings>(`bugBashProSettings_${VSS.getWebContext().project.id}`, settings, false);
            
            SettingsActionsHub.UpdateBugBashSettings.invoke(updatedSettings);
            StoresHub.bugBashSettingsStore.setLoading(false);
        }
        catch (e) {
            StoresHub.bugBashSettingsStore.setLoading(false);
            throw e.message;
        }
    }

    export async function initializeUserSettings() {
        if (StoresHub.userSettingsStore.isLoaded()) {
            SettingsActionsHub.InitializeUserSettings.invoke(null);
        }
        else if (!StoresHub.userSettingsStore.isLoading()) {
            StoresHub.userSettingsStore.setLoading(true);
            const settings = await ExtensionDataManager.readDocuments<UserSettings>(`bugBashProUserSettings_${VSS.getWebContext().project.id}`, false);
            SettingsActionsHub.InitializeUserSettings.invoke(settings);
            StoresHub.userSettingsStore.setLoading(false);
        }
    }

    export async function updateUserSettings(userId: string, settings: UserSettings) {
        if (!StoresHub.userSettingsStore.isLoading()) {
            StoresHub.userSettingsStore.setLoading(true);
            const settings = await ExtensionDataManager.addOrUpdateDocument<UserSettings>(`bugBashProUserSettings_${VSS.getWebContext().project.id}`, false);
            SettingsActionsHub.InitializeUserSettings.invoke(settings);
            StoresHub.userSettingsStore.setLoading(false);
        }
    }
}