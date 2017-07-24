import { ExtensionDataManager } from "VSTS_Extension/Utilities/ExtensionDataManager";

import { StoresHub } from "../Stores/StoresHub";
import { BugBashActionsHub } from "./ActionsHub";
import { IBugBash } from "../Interfaces";

export module BugBashActions {
    export async function initializeAllBugBashes() {
        if (StoresHub.bugBashStore.isLoaded()) {
            BugBashActionsHub.InitializeAllBugBashes.invoke(null);
        }
        else if (!StoresHub.bugBashStore.isLoading()) {
            StoresHub.bugBashStore.setLoading(true);
            let bugBashes = await ExtensionDataManager.readDocuments<IBugBash>("bugbashes", false);
            for(let bugBash of bugBashes) {
                translateDates(bugBash);
            }

            BugBashActionsHub.InitializeAllBugBashes.invoke(bugBashes);
            StoresHub.bugBashStore.setLoading(false);
        }
    } 

    export async function initializeBugBash(bugBashId: string) {
        if (StoresHub.bugBashStore.isLoaded(bugBashId)) {
            BugBashActionsHub.InitializeBugBash.invoke(null);
        }
        else if (!StoresHub.bugBashStore.isLoading(bugBashId)) {
            StoresHub.bugBashStore.setLoading(true, bugBashId);
            let bugBash = await ExtensionDataManager.readDocument<IBugBash>("bugbashes", bugBashId, null, false);
            if (bugBash) {
                translateDates(bugBash);
                BugBashActionsHub.InitializeBugBash.invoke(bugBash);
                StoresHub.bugBashStore.setLoading(false, bugBashId);
            }
            else {
                StoresHub.bugBashStore.setLoading(false, bugBashId);
                throw "This instance of bug bash does not exist.";
            }
        }
    } 

    export async function refreshBugBash(bugBashId: string) {
        if (!StoresHub.bugBashStore.isLoading(bugBashId)) {
            StoresHub.bugBashStore.setLoading(true, bugBashId);

            let bugBash = await ExtensionDataManager.readDocument<IBugBash>("bugbashes", bugBashId, null, false);
            if (bugBash) {
                translateDates(bugBash);
                BugBashActionsHub.RefreshBugBash.invoke(bugBash);
                StoresHub.bugBashStore.setLoading(false, bugBashId);
            }
            else {
                BugBashActionsHub.DeleteBugBash.invoke(bugBashId);
                StoresHub.bugBashStore.setLoading(false, bugBashId);
                throw "This instance of bug bash does not exist.";
            }
        }
    } 

    export async function refreshAllBugBashes() {
        if (!StoresHub.bugBashStore.isLoading()) {
            StoresHub.bugBashStore.setLoading(true);

            let bugBashes = await ExtensionDataManager.readDocuments<IBugBash>("bugbashes", false);
            for(let bugBash of bugBashes) {
                translateDates(bugBash);
            }
            
            BugBashActionsHub.RefreshAllBugBashes.invoke(bugBashes);
            StoresHub.bugBashStore.setLoading(false);
        }
    }

    export async function updateBugBash(bugBash: IBugBash) {
        if (!StoresHub.bugBashStore.isLoading(bugBash.id)) {
            StoresHub.bugBashStore.setLoading(true, bugBash.id);

            try {
                let savedBugBash = await ExtensionDataManager.updateDocument<IBugBash>("bugbashes", bugBash, false);
                translateDates(savedBugBash);
                
                BugBashActionsHub.UpdateBugBash.invoke(savedBugBash);
                StoresHub.bugBashStore.setLoading(false, bugBash.id);

                return savedBugBash;
            }
            catch (e) {
                StoresHub.bugBashStore.setLoading(false, bugBash.id);
                throw "This bug bash instance has been modified by some one else. Please refresh the instance to get the latest version and try updating it again.";
            }
        }
    }

    export async function createBugBash(bugBash: IBugBash): Promise<IBugBash> {
        if (!StoresHub.bugBashStore.isLoading()) {
            StoresHub.bugBashStore.setLoading(true);

            try {
                let cloneBugBash = {...bugBash};
                cloneBugBash.id = Date.now().toString();

                const savedBugBash = await ExtensionDataManager.createDocument<IBugBash>("bugbashes", cloneBugBash, false);
                translateDates(savedBugBash);            
                
                BugBashActionsHub.CreateBugBash.invoke(savedBugBash);
                StoresHub.bugBashStore.setLoading(false);
                return savedBugBash;
            }
            catch (e) {
                StoresHub.bugBashStore.setLoading(false);
                throw e.message;
            }
        }
    }

    export async function deleteBugBash(bugBashId: string) {
        if (!StoresHub.bugBashStore.isLoading(bugBashId)) {
            StoresHub.bugBashStore.setLoading(true, bugBashId);

            try {            
                await ExtensionDataManager.deleteDocument<IBugBash>("bugbashes", bugBashId, false);                
            }
            catch (e) {
                // eat exception
            }

            BugBashActionsHub.DeleteBugBash.invoke(bugBashId);
            StoresHub.bugBashStore.setLoading(false, bugBashId);
        }
    }

    function translateDates(bugBash: IBugBash) {
        if (typeof bugBash.startTime === "string") {
            if ((bugBash.startTime as string).trim() === "") {
                bugBash.startTime = undefined;
            }
            else {
                bugBash.startTime = new Date(bugBash.startTime);
            }
        }
        if (typeof bugBash.endTime === "string") {
            if ((bugBash.endTime as string).trim() === "") {
                bugBash.endTime = undefined;
            }
            else {
                bugBash.endTime = new Date(bugBash.endTime);
            }
        }
    }  
}