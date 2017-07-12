import { ExtensionDataManager } from "VSTS_Extension/Utilities/ExtensionDataManager";
import { parseUniquefiedIdentityName } from "VSTS_Extension/Components/WorkItemControls/IdentityView";
import { WorkItemTemplateItemActions } from "VSTS_Extension/Flux/Actions/WorkItemTemplateItemActions";
import { TeamFieldActions } from "VSTS_Extension/Flux/Actions/TeamFieldActions";

import Context = require("VSS/Context");
import Utils_Date = require("VSS/Utils/Date");
import { WorkItem } from "TFS/WorkItemTracking/Contracts";
import { UrlActions } from "../Constants";

import { BugBashItemActionsCreator } from "./ActionsCreator";
import { StoresHub } from "../Stores/StoresHub";
import { BugBashItemStore } from "../Stores/BugBashItemStore";
import { BugBashStore } from "../Stores/BugBashStore";
import { BugBashItemCommentStore } from "../Stores/BugBashItemCommentStore";
import { IBugBashItem, IBugBash, IBugBashItemComment, IAcceptedItemViewModel } from "../Interfaces";
import { createWorkItem, updateWorkItem, BugBashItemHelpers } from "../Helpers";

export module BugBashItemCommentActions {
    export async function initializeItems(bugBashId: string) {
        if (StoresHub.bugBashItemStore.isLoaded(bugBashId)) {
            BugBashItemActionsCreator.InitializeItems.invoke(null);
        }
        else if (!StoresHub.bugBashItemStore.isLoading(bugBashId)) {
            StoresHub.bugBashItemStore.setLoading(true, bugBashId);

            const items = await ExtensionDataManager.readDocuments<IBugBashItem>(getBugBashCollectionKey(bugBashId), false);
            for(let item of items) {
                translateDates(item);
            }
            
            BugBashItemActionsCreator.InitializeItems.invoke({bugBashId: bugBashId, bugBashItems: items});
            StoresHub.bugBashItemStore.setLoading(false, bugBashId);
        }
    } 

    export async function refreshItems(bugBashId: string) {
        if (!StoresHub.bugBashItemStore.isLoading(bugBashId)) {
            StoresHub.bugBashItemStore.setLoading(true, bugBashId);

            const items = await ExtensionDataManager.readDocuments<IBugBashItem>(getBugBashCollectionKey(bugBashId), false);
            for(let item of items) {
                translateDates(item);
            }
            
            BugBashItemActionsCreator.RefreshItems.invoke({bugBashId: bugBashId, bugBashItems: items});
            StoresHub.bugBashItemStore.setLoading(false, bugBashId);
        }
    } 

    export async function refreshItem(bugBashId: string, bugBashItemId: string) {
        if (!StoresHub.bugBashItemStore.isLoading(bugBashId) && !StoresHub.bugBashItemStore.isLoading(bugBashItemId)) {
            StoresHub.bugBashItemStore.setLoading(true, bugBashItemId);
            const item = await ExtensionDataManager.readDocument<IBugBashItem>(getBugBashCollectionKey(bugBashId), bugBashItemId, null, false);

            if (item) {
                translateDates(item);
                
                BugBashItemActionsCreator.RefreshItem.invoke({bugBashId: bugBashId, bugBashItem: item});
                StoresHub.bugBashItemStore.setLoading(false, bugBashItemId);
            }
            else {
                StoresHub.bugBashItemStore.setLoading(false, bugBashItemId);
                throw "This instance of bug bash item does not exist.";
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
        if (!StoresHub.bugBashItemStore.isLoading(bugBashId) && !StoresHub.bugBashItemStore.isLoading(bugBashItem.id)) {
            StoresHub.bugBashItemStore.setLoading(true, bugBashItem.id);

            try {
                let savedBugBashItem = await ExtensionDataManager.updateDocument(getBugBashCollectionKey(bugBashId), bugBashItem, false);
                translateDates(savedBugBashItem);
                
                BugBashItemActionsCreator.UpdateItem.invoke({bugBashId: bugBashId, bugBashItem: savedBugBashItem});
                StoresHub.bugBashItemStore.setLoading(false, bugBashItem.id);
            }
            catch (e) {
                StoresHub.bugBashItemStore.setLoading(false, bugBashItem.id);
                throw "This bug bash item has been modified by some one else. Please refresh the item to get the latest version and try updating it again.";
            }
        }
    }

    export async function createBugBashItem(bugBashId: string, bugBashItem: IBugBashItem) {
        if (!StoresHub.bugBashItemStore.isLoading(bugBashId)) {
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
                throw e.message;
            }
        }
    }
    
    export async function acceptBugBashItem(bugBashId: string, bugBashItem: IBugBashItem) {
        if (!StoresHub.bugBashItemStore.isLoading(bugBashId) && !StoresHub.bugBashItemStore.isLoading(bugBashItem.id)) {
            StoresHub.bugBashItemStore.setLoading(true, bugBashItem.id);

            try {
                let savedBugBashItem = await acceptItem(bugBashItem);
                
                BugBashItemActionsCreator.AcceptItem.invoke({bugBashId: bugBashId, bugBashItem: savedBugBashItem.model});
                StoresHub.bugBashItemStore.setLoading(false, bugBashItem.id);
            }
            catch (e) {
                StoresHub.bugBashItemStore.setLoading(false, bugBashItem.id);
                throw e;
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

    async function acceptItem(model: IBugBashItem): Promise<IAcceptedItemViewModel> {
        let updatedItem: IBugBashItem;
        let savedWorkItem: WorkItem;
        const bugBash = StoresHub.bugBashStore.getItem(model.bugBashId);

        // read bug bash wit template
        try {
            await WorkItemTemplateItemActions.initializeWorkItemTemplateItem(bugBash.acceptTemplate.team, bugBash.acceptTemplate.templateId);
        }
        catch (e) {
            throw `Bug bash template '${bugBash.acceptTemplate.templateId}' does not exist in team '${bugBash.acceptTemplate.team}'`;
        }

        try {
            await TeamFieldActions.initializeTeamFields(updatedItem.teamId);
        }
        catch (e) {
            throw `Cannot read team field value for team: ${updatedItem.teamId}.`;
        }

        try {
            // first do a empty save to check if its the latest version of the item
            updatedItem = await ExtensionDataManager.updateDocument(getBugBashCollectionKey(model.bugBashId), model, false);
        }
        catch (e) {
            throw "This item has been modified by some one else. Please refresh the item to get the latest version and try updating it again.";
        }

        const template = StoresHub.workItemTemplateItemStore.getItem(bugBash.acceptTemplate.templateId);
        const teamFieldValue = StoresHub.teamFieldStore.getItem(updatedItem.teamId);
        let fieldValues = {...template.fields};
        fieldValues["System.Title"] = updatedItem.title;
        fieldValues[bugBash.itemDescriptionField] = updatedItem.description;            
        fieldValues[teamFieldValue.field.referenceName] = teamFieldValue.defaultValue;        

        if (fieldValues["System.Tags-Add"]) {
            fieldValues["System.Tags"] = fieldValues["System.Tags-Add"];
        }            

        delete fieldValues["System.Tags-Add"];
        delete fieldValues["System.Tags-Remove"];

        try {
            // create work item
            savedWorkItem = await createWorkItem(bugBash.workItemType, fieldValues);
        }
        catch (e) {
            BugBashItemActionsCreator.UpdateItem.invoke({bugBashId: updatedItem.bugBashId, bugBashItem: updatedItem});
            throw e.message;
        }
        
        // associate work item with bug bash item
        updatedItem.workItemId = savedWorkItem.id;
        updatedItem.rejected = false;
        updatedItem.rejectedBy = "";
        updatedItem.rejectReason = "";
        updatedItem = await ExtensionDataManager.updateDocument(getBugBashCollectionKey(updatedItem.bugBashId), updatedItem, false);

        addExtraFieldsToWorkitem(savedWorkItem.id, updatedItem);

        return {
            model: updatedItem,
            workItem: savedWorkItem
        };
    }

    function addExtraFieldsToWorkitem(workItemId: number, model: IBugBashItem) {
        let fieldValues: IDictionaryStringTo<string> = {};
        const bugBash = StoresHub.bugBashStore.getItem(model.bugBashId);

        fieldValues["System.History"] = getAcceptedItemComment(bugBash, model);

        updateWorkItem(workItemId, fieldValues);
    }

    function getAcceptedItemComment(bugBash: IBugBash, model: IBugBashItem): string {
        const pageContext = Context.getPageContext();
        const navigation = pageContext.navigation;
        const webContext = VSS.getWebContext();
        const bugBashUrl = `${webContext.collection.uri}/${webContext.project.name}/_${navigation.currentController}/${navigation.currentAction}/${navigation.currentParameters}#_a=${UrlActions.ACTION_VIEW}&id=${bugBash.id}`;
        
        const entity = parseUniquefiedIdentityName(model.createdBy);

        let commentToSave = `
            Created from <a href='${bugBashUrl}' target='_blank'>${bugBash.title}</a> bug bash on behalf of <a href='mailto:${entity.uniqueName || entity.displayName || ""}' data-vss-mention='version:1.0'>@${entity.displayName}</a>
        `;

        let discussionComments = StoresHub.bugBashItemCommentStore.getItem(model.id);

        if (discussionComments && discussionComments.length > 0) {
            discussionComments = discussionComments.slice();
            discussionComments = discussionComments.sort((c1: IBugBashItemComment, c2: IBugBashItemComment) => {
                return Utils_Date.defaultComparer(c1.createdDate, c2.createdDate);
            });

            commentToSave += "<div style='margin: 15px 0;font-size: 15px; font-weight: bold; text-decoration: underline;'>Discussions :</div>";

            for (const comment of discussionComments) {
                commentToSave += `
                    <div style='border-bottom: 1px solid #ccc; padding-bottom: 5px; margin-bottom: 10px;'>
                        <div><span style='font-size: 13px; font-weight: bold;'>${comment.createdBy}</span> wrote:</div>
                        <div>${comment.content}</div>
                    </div>
                `;
            }            
        }

        return commentToSave;
    }
}