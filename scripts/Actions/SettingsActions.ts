import { ExtensionDataManager } from "MB/Utils/ExtensionDataManager";

import { SettingsActionsHub, BugBashErrorMessageActionsHub } from "./ActionsHub";
import { IUserSettings, IBugBashSettings } from "../Interfaces";
import { StoresHub } from "../Stores/StoresHub";
import { ErrorKeys } from "../Constants";

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
            BugBashErrorMessageActionsHub.DismissErrorMessage.invoke(ErrorKeys.BugBashSettingsError);
        }
        catch (e) {
            BugBashErrorMessageActionsHub.PushErrorMessage.invoke({
                errorMessage: "Settings could not be saved due to an unknown error. Please refresh the page and try again.",
                errorKey: ErrorKeys.BugBashSettingsError
            });
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
            BugBashErrorMessageActionsHub.DismissErrorMessage.invoke(ErrorKeys.BugBashSettingsError);
        }
        catch (e) {
            BugBashErrorMessageActionsHub.PushErrorMessage.invoke({
                errorMessage: "Settings could not be saved due to an unknown error. Please refresh the page and try again.",
                errorKey: ErrorKeys.BugBashSettingsError
            });
        }
    }
}