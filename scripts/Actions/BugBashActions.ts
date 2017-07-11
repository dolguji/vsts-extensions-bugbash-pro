import { ExtensionDataManager } from "VSTS_Extension/Utilities/ExtensionDataManager";

import { StoresHub } from "../Stores/StoresHub";
import { BugBashActionsCreator } from "./ActionsCreator";
import { BugBashStore } from "../Stores/BugBashStore";
import { IBugBash } from "../Interfaces";

export module BugBashActions {
    export async function initializeAllBugBashes() {
        if (StoresHub.bugBashStore.isLoaded()) {
            BugBashActionsCreator.InitializeAllBugBashes.invoke(null);
        }
        else if (!StoresHub.bugBashStore.isLoading()) {
            StoresHub.bugBashStore.setLoading(true);

            try {
                let bugBashes = await ExtensionDataManager.readDocuments<IBugBash>("bugbashes", false);
                for(let bugBash of bugBashes) {
                    translateDates(bugBash);
                }

                BugBashActionsCreator.InitializeAllBugBashes.invoke(bugBashes);
                StoresHub.bugBashStore.setLoading(false);
            }
            catch (e) {
                StoresHub.bugBashStore.setLoading(false);
                throw e;                
            }
        }
    } 

    export async function initializeBugBash(bugBashId: string) {
        if (StoresHub.bugBashStore.isLoaded(bugBashId)) {
            BugBashActionsCreator.InitializeBugBash.invoke(null);
        }
        else if (!StoresHub.bugBashStore.isLoading(bugBashId)) {
            StoresHub.bugBashStore.setLoading(true, bugBashId);

            try {
                let bugBash = await ExtensionDataManager.readDocument<IBugBash>("bugbashes", bugBashId, null, false);
                if (bugBash) {
                    translateDates(bugBash);                    

                    BugBashActionsCreator.InitializeBugBash.invoke(bugBash);
                    StoresHub.bugBashStore.setLoading(false, bugBashId);
                }
                else {
                    StoresHub.bugBashStore.setLoading(false, bugBashId);
                    throw new Error("This instance of bug bash does not exist.");
                }          
            }
            catch (e) {
                StoresHub.bugBashStore.setLoading(false, bugBashId);
                throw e;
            }
        }
    } 

    export async function refreshAllBugBashes() {
        if (!StoresHub.bugBashStore.isLoading()) {
            StoresHub.bugBashStore.setLoading(true);

            try {
                let bugBashes = await ExtensionDataManager.readDocuments<IBugBash>("bugbashes", false);
                for(let bugBash of bugBashes) {
                    translateDates(bugBash);
                }
                
                BugBashActionsCreator.RefreshAllBugBashes.invoke(bugBashes);
                StoresHub.bugBashStore.setLoading(false);
            }
            catch (e) {
                StoresHub.bugBashStore.setLoading(false);
                throw e;
            }
        }
    }

    export async function updateBugBash(bugBash: IBugBash) {
        if (!StoresHub.bugBashStore.isLoading(bugBash.id)) {
            StoresHub.bugBashStore.setLoading(true, bugBash.id);

            try {
                let savedBugBash = await ExtensionDataManager.updateDocument<IBugBash>("bugbashes", bugBash, false);
                translateDates(savedBugBash);
                
                BugBashActionsCreator.UpdateBugBash.invoke(savedBugBash);
                StoresHub.bugBashStore.setLoading(false, bugBash.id);
            }
            catch (e) {
                StoresHub.bugBashStore.setLoading(false, bugBash.id);
                throw new Error("This bug bash instance has been modified by some one else. Please refresh the page to get the latest version and try updating it again.");
            }
        }
    }

    export async function createBugBash(bugBash: IBugBash) {
        if (!StoresHub.bugBashStore.isLoading()) {
            StoresHub.bugBashStore.setLoading(true);

            try {
                let model = {...bugBash};
                model.id = Date.now().toString();

                const savedBugBash = await ExtensionDataManager.createDocument<IBugBash>("bugbashes", model, false);
                translateDates(savedBugBash);            
                
                BugBashActionsCreator.CreateBugBash.invoke(savedBugBash);
                StoresHub.bugBashStore.setLoading(false);
            }
            catch (e) {
                StoresHub.bugBashStore.setLoading(false);
                throw e;
            }
        }
    }

    export async function deleteBugBash(bugBash: IBugBash) {
        if (!StoresHub.bugBashStore.isLoading(bugBash.id)) {
            StoresHub.bugBashStore.setLoading(true, bugBash.id);

            try {            
                await ExtensionDataManager.deleteDocument<IBugBash>("bugbashes", bugBash.id, false);                
            }
            catch (e) {
                
            }
            finally {                
                BugBashActionsCreator.DeleteBugBash.invoke(bugBash);
                StoresHub.bugBashStore.setLoading(false, bugBash.id);
            }
        }
    }

    function translateDates(item: IBugBash) {
        if (typeof item.startTime === "string") {
            if ((item.startTime as string).trim() === "") {
                item.startTime = undefined;
            }
            else {
                item.startTime = new Date(item.startTime);
            }
        }
        if (typeof item.endTime === "string") {
            if ((item.endTime as string).trim() === "") {
                item.endTime = undefined;
            }
            else {
                item.endTime = new Date(item.endTime);
            }
        }
    }  
}