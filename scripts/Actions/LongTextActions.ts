import { ExtensionDataManager } from "MB/Utils/ExtensionDataManager";

import { StoresHub } from "../Stores/StoresHub";
import { LongTextActionsHub, BugBashErrorMessageActionsHub } from "./ActionsHub";
import { ErrorKeys } from "../Constants";
import { ILongText } from "../Interfaces";

export module LongTextActions {
    export function fireStoreChange() {
        LongTextActionsHub.FireStoreChange.invoke(null);
    }

    export async function initializeLongText(id: string) {
        if (StoresHub.longTextStore.isLoaded(id)) {
            LongTextActionsHub.AddOrUpdateLongText.invoke(null);
        }
        else if (!StoresHub.longTextStore.isLoading(id)) {
            StoresHub.longTextStore.setLoading(true, id);

            const longText = await ExtensionDataManager.readDocument<ILongText>("longtexts", id, {
                id: id,
                text: ""
            }, false);

            LongTextActionsHub.AddOrUpdateLongText.invoke(longText);
            StoresHub.longTextStore.setLoading(false, id);
        }
    } 

    export async function refreshLongText(id: string) {
        if (!StoresHub.longTextStore.isLoading(id)) {
            StoresHub.longTextStore.setLoading(true, id);

            const longText = await ExtensionDataManager.readDocument<ILongText>("longtexts", id, {
                id: id,
                text: ""
            }, false);

            LongTextActionsHub.AddOrUpdateLongText.invoke(longText);
            StoresHub.longTextStore.setLoading(false, id);
            BugBashErrorMessageActionsHub.DismissErrorMessage.invoke(ErrorKeys.BugBashDetailsError);
        }
    }

    export async function addOrUpdateLongText(longText: ILongText) {
        if (!StoresHub.longTextStore.isLoading(longText.id)) {
            StoresHub.longTextStore.setLoading(true, longText.id);
            try {
                const savedLongText = await ExtensionDataManager.addOrUpdateDocument<ILongText>("longtexts", longText, false);
                LongTextActionsHub.AddOrUpdateLongText.invoke(savedLongText);
                StoresHub.longTextStore.setLoading(false, longText.id);
                BugBashErrorMessageActionsHub.DismissErrorMessage.invoke(ErrorKeys.BugBashDetailsError);
            }
            catch (e) {
                StoresHub.longTextStore.setLoading(false, longText.id);
                BugBashErrorMessageActionsHub.PushErrorMessage.invoke({
                    errorMessage: "This text has been modified by some one else. Please refresh the instance to get the latest version and try updating it again.", 
                    errorKey: ErrorKeys.BugBashDetailsError
                });
            }
        }
    }
}