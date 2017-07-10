import { StoreFactory } from "VSTS_Extension/Flux/Stores/BaseStore";
import { ExtensionDataManager } from "VSTS_Extension/Utilities/ExtensionDataManager";

import { BugBashActionsCreator } from "./ActionsCreator";
import { BugBashStore } from "../Stores/BugBashStore";
import { IBugBash } from "../Interfaces";

export module BugBashActions {
    var bugBashStore: BugBashStore = StoreFactory.getInstance<BugBashStore>(BugBashStore);

    export async function initializeAllBugBashes() {
        if (bugBashStore.isLoaded()) {
            BugBashActionsCreator.InitializeAllBugBashes.invoke(null);
        }
        else if (!bugBashStore.isLoading()) {
            bugBashStore.setLoading(true);

            try {
                let bugBashes = await ExtensionDataManager.readDocuments<IBugBash>("bugbashes", false);
                for(let bugBash of bugBashes) {
                    translateDates(bugBash);
                }

                bugBashStore.setLoading(false);
                bugBashStore.setError(null);

                BugBashActionsCreator.InitializeAllBugBashes.invoke(bugBashes);
            }
            catch (e) {
                bugBashStore.setLoading(false);
                bugBashStore.setError(e.message || e);
            }
        }
    } 

    export async function initializeBugBash(bugBashId: string) {
        if (bugBashStore.isLoaded(bugBashId)) {
            BugBashActionsCreator.InitializeBugBash.invoke(null);
        }
        else if (!bugBashStore.isLoading(bugBashId)) {
            bugBashStore.setLoading(true, bugBashId);

            try {
                let bugBash = await ExtensionDataManager.readDocument<IBugBash>("bugbashes", bugBashId, null, false);
                if (bugBash) {
                    translateDates(bugBash);
                    bugBashStore.setLoading(false, bugBashId);
                    bugBashStore.setError(null, bugBashId);

                    BugBashActionsCreator.InitializeBugBash.invoke(bugBash);
                }
                else {
                    bugBashStore.setLoading(false, bugBashId);
                    bugBashStore.setError("This instance of bug bash does not exist.", bugBashId);
                }          
            }
            catch (e) {
                bugBashStore.setLoading(false, bugBashId);
                bugBashStore.setError(e.message || e, bugBashId);
            }
        }
    } 

    export async function refreshAllBugBashes() {
        if (!bugBashStore.isLoading()) {
            bugBashStore.setLoading(true);

            try {
                let bugBashes = await ExtensionDataManager.readDocuments<IBugBash>("bugbashes", false);
                for(let bugBash of bugBashes) {
                    translateDates(bugBash);
                }

                bugBashStore.setLoading(false);
                bugBashStore.setError(null);

                BugBashActionsCreator.RefreshAllBugBashes.invoke(bugBashes);
            }
            catch (e) {
                bugBashStore.setLoading(false);
                bugBashStore.setError(e.message || e);
            }
        }
    }

    export async function updateBugBash(bugBash: IBugBash) {
        if (!bugBashStore.isLoading(bugBash.id)) {
            bugBashStore.setLoading(true, bugBash.id);

            try {
                let savedBugBash = await ExtensionDataManager.updateDocument<IBugBash>("bugbashes", bugBash, false);
                translateDates(savedBugBash);

                bugBashStore.setLoading(false, bugBash.id);
                bugBashStore.setError(null, bugBash.id);

                BugBashActionsCreator.UpdateBugBash.invoke(savedBugBash);
            }
            catch (e) {
                bugBashStore.setLoading(false, bugBash.id);
                bugBashStore.setError(e.message || e, bugBash.id);
            }
        }
    }

    export async function createBugBash(bugBash: IBugBash) {
        try {
            let model = {...bugBash};
            model.id = Date.now().toString();

            const savedBugBash = await ExtensionDataManager.createDocument<IBugBash>("bugbashes", model, false);
            translateDates(savedBugBash);            

            BugBashActionsCreator.CreateBugBash.invoke(savedBugBash);
        }
        catch (e) {
            
        }
    }

    export async function deleteBugBash(bugBash: IBugBash) {
        if (!bugBashStore.isLoading(bugBash.id)) {
            bugBashStore.setLoading(true, bugBash.id);

            try {            
                await ExtensionDataManager.deleteDocument<IBugBash>("bugbashes", bugBash.id, false);
                bugBashStore.setLoading(false, bugBash.id);
                bugBashStore.setError(null, bugBash.id);
                BugBashActionsCreator.DeleteBugBash.invoke(bugBash);
            }
            catch (e) {
                bugBashStore.setLoading(false, bugBash.id);
                bugBashStore.setError("This instance of bug bash is already deleted.", bugBash.id);
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