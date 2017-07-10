import { StoreFactory } from "VSTS_Extension/Flux/Stores/BaseStore";
import { ExtensionDataManager } from "VSTS_Extension/Utilities/ExtensionDataManager";

import { BugBashItemActionsCreator } from "./ActionsCreator";
import { BugBashItemStore } from "../Stores/BugBashItemStore";
import { IBugBashItem } from "../Interfaces";
import { createWorkItem, updateWorkItem, BugBashItemHelpers } from "../Helpers";

export module BugBashItemCommentActions {
    var bugBashItemStore: BugBashItemStore = StoreFactory.getInstance<BugBashItemStore>(BugBashItemStore);

    export async function initializeItems(bugBashId: string) {
        if (bugBashItemStore.isLoaded(bugBashId)) {
            BugBashItemActionsCreator.InitializeItems.invoke(null);
        }
        else if (!bugBashItemStore.isLoading(bugBashId)) {
            bugBashItemStore.setLoading(true, bugBashId);

            try {
                const items = await ExtensionDataManager.readDocuments<IBugBashItem>(getBugBashCollectionKey(bugBashId), false);
                for(let item of items) {
                    translateDates(item);
                }
                bugBashItemStore.setLoading(false, bugBashId);
                bugBashItemStore.setError(null, bugBashId);

                BugBashItemActionsCreator.InitializeItems.invoke({bugBashId: bugBashId, bugBashItems: items});
            }
            catch (e) {
                bugBashItemStore.setLoading(false, bugBashId);
                bugBashItemStore.setError(e.message || e, bugBashId);
            }
        }
    } 

    export async function refreshItems(bugBashId: string) {
        if (!bugBashItemStore.isLoading(bugBashId)) {
            bugBashItemStore.setLoading(true, bugBashId);

            try {
                const items = await ExtensionDataManager.readDocuments<IBugBashItem>(getBugBashCollectionKey(bugBashId), false);
                for(let item of items) {
                    translateDates(item);
                }
                bugBashItemStore.setLoading(false, bugBashId);
                bugBashItemStore.setError(null, bugBashId);

                BugBashItemActionsCreator.RefreshItems.invoke({bugBashId: bugBashId, bugBashItems: items});
            }
            catch (e) {
                bugBashItemStore.setLoading(false, bugBashId);
                bugBashItemStore.setError(e.message || e, bugBashId);
            }
        }
    } 

    export async function refreshItem(bugBashId: string, bugBashItemId: string) {
        if (!bugBashItemStore.isLoading(bugBashId)) {
            bugBashItemStore.setLoading(true, bugBashId);

            try {
                const item = await ExtensionDataManager.readDocument<IBugBashItem>(getBugBashCollectionKey(bugBashId), bugBashItemId, null, false);

                if (item) {
                    translateDates(item);
                    bugBashItemStore.setLoading(false, bugBashId);
                    bugBashItemStore.setError(null, bugBashId);

                    BugBashItemActionsCreator.RefreshItem.invoke({bugBashId: bugBashId, bugBashItem: item});
                }
                else {
                    bugBashItemStore.setLoading(false, bugBashId);
                    bugBashItemStore.setError("This bug bash item does not exist anymore", bugBashId);
                }
            }
            catch (e) {
                bugBashItemStore.setLoading(false, bugBashId);
                bugBashItemStore.setError(e.message || e, bugBashId);
            }
        }
    } 

    export async function saveItem(bugBashId: string, bugBashItem: IBugBashItem) {
        const isNew = BugBashItemHelpers.isNew(bugBashItem);        

        if (isNew) {
            this.createBugBashItem(bugBashId, bugBashItem);
        }
        else {
            this.updateBugBashItem(bugBashId, bugBashItem);
        }
    }

    export async function updateBugBashItem(bugBashId: string, bugBashItem: IBugBashItem) {
        if (!bugBashItemStore.isLoading(bugBashId)) {
            bugBashItemStore.setLoading(true, bugBashId);

            try {
                let savedBugBashItem = await ExtensionDataManager.updateDocument(getBugBashCollectionKey(bugBashId), bugBashItem, false);
                translateDates(savedBugBashItem);

                bugBashItemStore.setLoading(false, bugBashId);
                bugBashItemStore.setError(null, bugBashId);

                BugBashItemActionsCreator.UpdateItem.invoke({bugBashId: bugBashId, bugBashItem: savedBugBashItem});
            }
            catch (e) {
                bugBashItemStore.setLoading(false, bugBashId);
                bugBashItemStore.setError(e.message || e, bugBashId);
            }
        }
    }

    export async function createBugBashItem(bugBashId: string, bugBashItem: IBugBashItem) {
        try {
            let cloneModel = BugBashItemHelpers.deepCopy(bugBashItem);

            cloneModel.id = `${bugBashId}_${Date.now().toString()}`;
            cloneModel.createdBy = `${VSS.getWebContext().user.name} <${VSS.getWebContext().user.uniqueName}>`;
            cloneModel.createdDate = new Date(Date.now());
            
            let savedBugBashItem = await ExtensionDataManager.createDocument(getBugBashCollectionKey(bugBashId), bugBashItem, false);
            translateDates(savedBugBashItem);         

            BugBashItemActionsCreator.CreateItem.invoke({bugBashId: bugBashId, bugBashItem: savedBugBashItem});
        }
        catch (e) {
            
        }
    }
    
    export async function acceptBugBashItem(bugBashId: string, bugBashItem: IBugBashItem) {
        if (!bugBashItemStore.isLoading(bugBashId)) {
            bugBashItemStore.setLoading(true, bugBashId);

            try {
                let savedBugBashItem = await ExtensionDataManager.updateDocument(getBugBashCollectionKey(bugBashId), bugBashItem, false);
                translateDates(savedBugBashItem);

                bugBashItemStore.setLoading(false, bugBashId);
                bugBashItemStore.setError(null, bugBashId);

                BugBashItemActionsCreator.AcceptItem.invoke({bugBashId: bugBashId, bugBashItem: savedBugBashItem});
            }
            catch (e) {
                bugBashItemStore.setLoading(false, bugBashId);
                bugBashItemStore.setError(e.message || e, bugBashId);
            }
        }
    }

    function getBugBashCollectionKey(bugBashId: string): string {
        return `BugBashCollection_${bugBashId}`;
    }

    function translateDates(model: IBugBashItem) {
        if (typeof model.createdDate === "string") {
            if ((model.createdDate as string).trim() === "") {
                model.createdDate = undefined;
            }
            else {
                model.createdDate = new Date(model.createdDate);
            }
        }

        model.teamId = model.teamId || "";
    }
}