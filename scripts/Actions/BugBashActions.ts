import { HostNavigationService } from 'VSS/SDK/Services/Navigation';
import { ExtensionDataManager } from "VSTS_Extension/Utilities/ExtensionDataManager";
import * as Utils_String from 'VSS/Utils/String';

import { StoresHub } from "../Stores/StoresHub";
import { BugBashActionsHub, BugBashErrorMessageActionsHub } from "./ActionsHub";
import { IBugBash } from "../Interfaces";
import { ErrorKeys, UrlActions } from "../Constants";

export module BugBashActions {
    export function fireStoreChange() {
        BugBashActionsHub.FireStoreChange.invoke(null);
    }

    export async function initializeAllBugBashes() {
        if (StoresHub.bugBashStore.isLoaded()) {
            BugBashActionsHub.InitializeAllBugBashes.invoke(null);
        }
        else if (!StoresHub.bugBashStore.isLoading()) {
            StoresHub.bugBashStore.setLoading(true);
            let bugBashModels = await ExtensionDataManager.readDocuments<IBugBash>("bugbashes", false);
            bugBashModels = bugBashModels.filter(b => Utils_String.equals(VSS.getWebContext().project.id, b.projectId, true));

            for (let bugBashModel of bugBashModels) {
                preProcessModel(bugBashModel);
            }

            BugBashActionsHub.InitializeAllBugBashes.invoke(bugBashModels);
            StoresHub.bugBashStore.setLoading(false);
        }
    } 

    export async function initializeBugBash(bugBashId: string) {
        if (StoresHub.bugBashStore.isLoaded(bugBashId)) {
            BugBashActionsHub.InitializeBugBash.invoke(null);
        }
        else if (!StoresHub.bugBashStore.isLoading(bugBashId)) {
            let error = false;
            StoresHub.bugBashStore.setLoading(true, bugBashId);
            let bugBashModel = await ExtensionDataManager.readDocument<IBugBash>("bugbashes", bugBashId, null, false);

            if (bugBashModel && Utils_String.equals(VSS.getWebContext().project.id, bugBashModel.projectId, true)) {
                preProcessModel(bugBashModel);
                BugBashActionsHub.InitializeBugBash.invoke(bugBashModel);
                StoresHub.bugBashStore.setLoading(false, bugBashId);

                BugBashErrorMessageActionsHub.DismissErrorMessage.invoke(ErrorKeys.BugBashError);
            }
            else if (bugBashModel && !Utils_String.equals(VSS.getWebContext().project.id, bugBashModel.projectId, true)) {
                StoresHub.bugBashStore.setLoading(false, bugBashId);
                BugBashErrorMessageActionsHub.PushErrorMessage.invoke({
                    errorMessage: `Bug Bash "${bugBashId}" is out of scope of current project.`, 
                    errorKey: ErrorKeys.DirectoryPageError
                });

                error = true;
            }
            else {
                StoresHub.bugBashStore.setLoading(false, bugBashId);
                BugBashErrorMessageActionsHub.PushErrorMessage.invoke({
                    errorMessage: `Bug Bash "${bugBashId}" does not exist.`,
                    errorKey: ErrorKeys.DirectoryPageError
                });

                error = true;
            }

            if (error) {
                const navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
                navigationService.updateHistoryEntry(UrlActions.ACTION_ALL, null, true);
            }
        }
    } 

    export async function refreshBugBash(bugBashId: string) {
        if (!StoresHub.bugBashStore.isLoading(bugBashId)) {
            let error = false;
            StoresHub.bugBashStore.setLoading(true, bugBashId);

            let bugBashModel = await ExtensionDataManager.readDocument<IBugBash>("bugbashes", bugBashId, null, false);

            if (bugBashModel && Utils_String.equals(VSS.getWebContext().project.id, bugBashModel.projectId, true)) {
                preProcessModel(bugBashModel);
                BugBashActionsHub.RefreshBugBash.invoke(bugBashModel);
                StoresHub.bugBashStore.setLoading(false, bugBashId);

                BugBashErrorMessageActionsHub.DismissErrorMessage.invoke(ErrorKeys.BugBashError);
            }
            else if (bugBashModel && !Utils_String.equals(VSS.getWebContext().project.id, bugBashModel.projectId, true)) {
                StoresHub.bugBashStore.setLoading(false, bugBashId);
                BugBashErrorMessageActionsHub.PushErrorMessage.invoke({
                    errorMessage: `Bug Bash "${bugBashId}" is out of scope of current project.`, 
                    errorKey: ErrorKeys.DirectoryPageError
                });

                error = true;
            }
            else {
                BugBashActionsHub.DeleteBugBash.invoke(bugBashId);
                StoresHub.bugBashStore.setLoading(false, bugBashId);
                BugBashErrorMessageActionsHub.PushErrorMessage.invoke({
                    errorMessage: `Bug Bash "${bugBashId}" does not exist.`,
                    errorKey: ErrorKeys.DirectoryPageError
                });

                error = true;
            }

            if (error) {
                const navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
                navigationService.updateHistoryEntry(UrlActions.ACTION_ALL, null, true);
            }
        }
    } 

    export async function refreshAllBugBashes() {
        if (!StoresHub.bugBashStore.isLoading()) {
            StoresHub.bugBashStore.setLoading(true);

            let bugBashModels = await ExtensionDataManager.readDocuments<IBugBash>("bugbashes", false);
            bugBashModels = bugBashModels.filter(b => Utils_String.equals(VSS.getWebContext().project.id, b.projectId, true));

            for(let bugBashModel of bugBashModels) {
                preProcessModel(bugBashModel);
            }
            
            BugBashActionsHub.RefreshAllBugBashes.invoke(bugBashModels);
            StoresHub.bugBashStore.setLoading(false);

            BugBashErrorMessageActionsHub.DismissErrorMessage.invoke(ErrorKeys.DirectoryPageError);
        }
    }

    export async function updateBugBash(bugBashModel: IBugBash) {
        if (!StoresHub.bugBashStore.isLoading(bugBashModel.id)) {
            StoresHub.bugBashStore.setLoading(true, bugBashModel.id);

            try {
                let updatedBugBashModel = await ExtensionDataManager.updateDocument<IBugBash>("bugbashes", bugBashModel, false);
                preProcessModel(updatedBugBashModel);
                
                BugBashActionsHub.UpdateBugBash.invoke(updatedBugBashModel);
                StoresHub.bugBashStore.setLoading(false, bugBashModel.id);

                BugBashErrorMessageActionsHub.DismissErrorMessage.invoke(ErrorKeys.BugBashError);
            }
            catch (e) {
                StoresHub.bugBashStore.setLoading(false, bugBashModel.id);
                BugBashErrorMessageActionsHub.PushErrorMessage.invoke({
                    errorMessage: "This bug bash instance has been modified by some one else. Please refresh the instance to get the latest version and try updating it again.", 
                    errorKey: ErrorKeys.BugBashError
                });
            }
        }
    }

    export async function createBugBash(bugBashModel: IBugBash) {
        if (!StoresHub.bugBashStore.isLoading()) {
            StoresHub.bugBashStore.setLoading(true);

            try {
                let cloneBugBashModel = {...bugBashModel};
                cloneBugBashModel.id = Date.now().toString();

                const createdBugBashModel = await ExtensionDataManager.createDocument<IBugBash>("bugbashes", cloneBugBashModel, false);
                preProcessModel(createdBugBashModel);            
                
                BugBashActionsHub.CreateBugBash.invoke(createdBugBashModel);
                StoresHub.bugBashStore.setLoading(false);
                
                BugBashErrorMessageActionsHub.DismissErrorMessage.invoke(ErrorKeys.BugBashError);
                
                const navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
                navigationService.updateHistoryEntry(UrlActions.ACTION_EDIT, { id: createdBugBashModel.id }, true);
            }
            catch (e) {
                StoresHub.bugBashStore.setLoading(false);
                BugBashErrorMessageActionsHub.PushErrorMessage.invoke({
                    errorMessage: e.message, 
                    errorKey: ErrorKeys.BugBashError
                });
            }
        }
    }

    export async function deleteBugBash(bugBashId: string) {
        if (!StoresHub.bugBashStore.isLoading(bugBashId)) {
            StoresHub.bugBashStore.setLoading(true, bugBashId);

            try {            
                await ExtensionDataManager.deleteDocument("bugbashes", bugBashId, false);                
            }
            catch (e) {
                // eat exception
            }

            BugBashActionsHub.DeleteBugBash.invoke(bugBashId);
            StoresHub.bugBashStore.setLoading(false, bugBashId);
        }
    }    

    function preProcessModel(bugBashModel: IBugBash) {
        if (typeof bugBashModel.startTime === "string") {
            if ((bugBashModel.startTime as string).trim() === "") {
                bugBashModel.startTime = undefined;
            }
            else {
                bugBashModel.startTime = new Date(bugBashModel.startTime);
            }
        }
        if (typeof bugBashModel.endTime === "string") {
            if ((bugBashModel.endTime as string).trim() === "") {
                bugBashModel.endTime = undefined;
            }
            else {
                bugBashModel.endTime = new Date(bugBashModel.endTime);
            }
        }

        // convert old format of accept template to new one
        if (bugBashModel["acceptTemplate"] != null) {
            bugBashModel.acceptTemplateId = bugBashModel["acceptTemplate"].templateId;
            bugBashModel.acceptTemplateTeam = bugBashModel["acceptTemplate"].team;

            delete bugBashModel["acceptTemplate"];
        }
    }  
}