import { ExtensionDataManager } from "VSTS_Extension/Utilities/ExtensionDataManager";
import { parseUniquefiedIdentityName } from "VSTS_Extension/Components/WorkItemControls/IdentityView";
import { WorkItemTemplateItemActions } from "VSTS_Extension/Flux/Actions/WorkItemTemplateItemActions";
import { TeamFieldActions } from "VSTS_Extension/Flux/Actions/TeamFieldActions";
import { WorkItemActions } from "VSTS_Extension/Flux/Actions/WorkItemActions";

import Context = require("VSS/Context");
import Utils_Date = require("VSS/Utils/Date");
import { WorkItem, WorkItemTemplate } from "TFS/WorkItemTracking/Contracts";

import { UrlActions } from "../Constants";
import { BugBashItemActionsHub } from "./ActionsHub";
import { StoresHub } from "../Stores/StoresHub";
import { IBugBashItem, IBugBashItemComment } from "../Interfaces";
import { BugBashItemHelpers } from "../Helpers";
import { BugBash } from "../ViewModels/BugBash";

export module BugBashItemActions {
    export async function initializeItems(bugBashId: string) {
        if (StoresHub.bugBashItemStore.isLoaded(bugBashId)) {
            BugBashItemActionsHub.InitializeBugBashItems.invoke(null);
        }
        else if (!StoresHub.bugBashItemStore.isLoading(bugBashId)) {
            StoresHub.bugBashItemStore.setLoading(true, bugBashId);

            const bugBashItems = await ExtensionDataManager.readDocuments<IBugBashItem>(getBugBashCollectionKey(bugBashId), false);
            for(let bugBashItem of bugBashItems) {
                translateDates(bugBashItem);
            }
            
            BugBashItemActionsHub.InitializeBugBashItems.invoke({bugBashId: bugBashId, bugBashItems: bugBashItems});
            StoresHub.bugBashItemStore.setLoading(false, bugBashId);
        }
    } 

    export async function refreshItems(bugBashId: string) {
        if (!StoresHub.bugBashItemStore.isLoading(bugBashId)) {
            StoresHub.bugBashItemStore.setLoading(true, bugBashId);

            const bugBashItems = await ExtensionDataManager.readDocuments<IBugBashItem>(getBugBashCollectionKey(bugBashId), false);
            for(let bugBashItem of bugBashItems) {
                translateDates(bugBashItem);
            }
            
            BugBashItemActionsHub.RefreshBugBashItems.invoke({bugBashId: bugBashId, bugBashItems: bugBashItems});
            StoresHub.bugBashItemStore.setLoading(false, bugBashId);
        }
    } 

    export async function refreshItem(bugBashId: string, bugBashItemId: string): Promise<IBugBashItem> {
        if (!StoresHub.bugBashItemStore.isLoading(bugBashId) && !StoresHub.bugBashItemStore.isLoading(bugBashItemId)) {
            StoresHub.bugBashItemStore.setLoading(true, bugBashItemId);
            const bugBashItem = await ExtensionDataManager.readDocument<IBugBashItem>(getBugBashCollectionKey(bugBashId), bugBashItemId, null, false);

            if (bugBashItem) {
                translateDates(bugBashItem);
                
                BugBashItemActionsHub.RefreshBugBashItem.invoke({bugBashId: bugBashId, bugBashItem: bugBashItem});
                StoresHub.bugBashItemStore.setLoading(false, bugBashItemId);

                return bugBashItem;
            }
            else {
                StoresHub.bugBashItemStore.setLoading(false, bugBashItemId);
                throw "This instance of bug bash item does not exist.";
            }
        }

        return null;
    } 

    export async function saveItem(bugBashId: string, bugBashItem: IBugBashItem): Promise<IBugBashItem> {
        const isNew = BugBashItemHelpers.isNew(bugBashItem);        

        if (isNew) {
            return this.createBugBashItem(bugBashId, bugBashItem);
        }
        else {
            return this.updateBugBashItem(bugBashId, bugBashItem);
        }
    }

    export async function updateBugBashItem(bugBashId: string, bugBashItem: IBugBashItem): Promise<IBugBashItem> {
        if (!StoresHub.bugBashItemStore.isLoading(bugBashId) && !StoresHub.bugBashItemStore.isLoading(bugBashItem.id)) {
            StoresHub.bugBashItemStore.setLoading(true, bugBashItem.id);

            try {
                let savedBugBashItem = await ExtensionDataManager.updateDocument(getBugBashCollectionKey(bugBashId), bugBashItem, false);
                translateDates(savedBugBashItem);
                
                BugBashItemActionsHub.UpdateBugBashItem.invoke({bugBashId: bugBashId, bugBashItem: savedBugBashItem});
                StoresHub.bugBashItemStore.setLoading(false, bugBashItem.id);

                return savedBugBashItem;
            }
            catch (e) {
                StoresHub.bugBashItemStore.setLoading(false, bugBashItem.id);
                throw "This bug bash item has been modified by some one else. Please refresh the item to get the latest version and try updating it again.";
            }
        }

        return null;
    }

    export async function createBugBashItem(bugBashId: string, bugBashItem: IBugBashItem): Promise<IBugBashItem> {
        if (!StoresHub.bugBashItemStore.isLoading(bugBashId)) {
            try {
                let cloneBugBashItem = {...bugBashItem};
                cloneBugBashItem.id = `${bugBashId}_${Date.now().toString()}`;
                cloneBugBashItem.createdBy = `${VSS.getWebContext().user.name} <${VSS.getWebContext().user.uniqueName}>`;
                cloneBugBashItem.createdDate = new Date(Date.now());
                
                let savedBugBashItem = await ExtensionDataManager.createDocument(getBugBashCollectionKey(bugBashId), cloneBugBashItem, false);
                translateDates(savedBugBashItem);         
                
                BugBashItemActionsHub.CreateBugBashItem.invoke({bugBashId: bugBashId, bugBashItem: savedBugBashItem});

                return savedBugBashItem;
            }
            catch (e) {
                throw e.message;
            }
        }

        return null;
    }

    export async function deleteBugBashItem(bugBashId: string, bugBashItemId: string) {
        if (!StoresHub.bugBashItemStore.isLoading(bugBashId) && !StoresHub.bugBashItemStore.isLoading(bugBashItemId)) {
            StoresHub.bugBashItemStore.setLoading(true, bugBashItemId);

            try {            
                await ExtensionDataManager.deleteDocument(getBugBashCollectionKey(bugBashId), bugBashItemId, false);                
            }
            catch (e) {
                // eat exception
            }

            BugBashItemActionsHub.DeleteBugBashItem.invoke({bugBashId: bugBashId, bugBashItemId: bugBashItemId});
            StoresHub.bugBashItemStore.setLoading(false, bugBashItemId);
        }
    }
    
    export async function acceptBugBashItem(bugBashId: string, bugBashItem: IBugBashItem): Promise<IBugBashItem> {
        if (!StoresHub.bugBashItemStore.isLoading(bugBashId) && !StoresHub.bugBashItemStore.isLoading(bugBashItem.id)) {
            StoresHub.bugBashItemStore.setLoading(true, bugBashItem.id);

            try {
                let acceptedBugBashItem = await acceptItem(bugBashItem);
                
                BugBashItemActionsHub.AcceptBugBashItem.invoke({bugBashId: bugBashId, bugBashItem: acceptedBugBashItem});
                StoresHub.bugBashItemStore.setLoading(false, bugBashItem.id);

                return acceptedBugBashItem;
            }
            catch (e) {
                StoresHub.bugBashItemStore.setLoading(false, bugBashItem.id);
                throw e;
            }
        }

        return null;
    }

    function getBugBashCollectionKey(bugBashId: string): string {
        return `BugBashCollection_${bugBashId}`;
    }

    function translateDates(bugBashItem: IBugBashItem) {
        if (typeof bugBashItem.createdDate === "string") {
            if ((bugBashItem.createdDate as string).trim() === "") {
                bugBashItem.createdDate = undefined;
            }
            else {
                bugBashItem.createdDate = new Date(bugBashItem.createdDate);
            }
        }

        bugBashItem.teamId = bugBashItem.teamId || "";
    }    

    async function acceptItem(bugBashItem: IBugBashItem): Promise<IBugBashItem> {
        let updatedBugBashItem: IBugBashItem;
        let savedWorkItem: WorkItem;
        const bugBash = StoresHub.bugBashStore.getItem(bugBashItem.bugBashId);
        let acceptTemplate: WorkItemTemplate;

        // read bug bash wit template
        if (bugBash.originalModel.acceptTemplate && bugBash.originalModel.acceptTemplate.team && bugBash.originalModel.acceptTemplate.templateId) {
            try {
                await WorkItemTemplateItemActions.initializeWorkItemTemplateItem(bugBash.originalModel.acceptTemplate.team, bugBash.originalModel.acceptTemplate.templateId);
                acceptTemplate = StoresHub.workItemTemplateItemStore.getItem(bugBash.originalModel.acceptTemplate.templateId);
            }
            catch (e) {
                throw `Bug bash template '${bugBash.originalModel.acceptTemplate.templateId}' does not exist in team '${bugBash.originalModel.acceptTemplate.team}'`;
            }
        }

        try {
            await TeamFieldActions.initializeTeamFields(bugBashItem.teamId);
        }
        catch (e) {
            throw `Cannot read team field value for team: ${bugBashItem.teamId}.`;
        }

        try {
            // first do a empty save to check if its the latest version of the item
            updatedBugBashItem = await ExtensionDataManager.updateDocument(getBugBashCollectionKey(bugBashItem.bugBashId), bugBashItem, false);
        }
        catch (e) {
            throw "This item has been modified by some one else. Please refresh the item to get the latest version and try updating it again.";
        }

        const teamFieldValue = StoresHub.teamFieldStore.getItem(updatedBugBashItem.teamId);
        let fieldValues = acceptTemplate ? {...acceptTemplate.fields} : {};
        fieldValues["System.Title"] = updatedBugBashItem.title;
        fieldValues[bugBash.originalModel.itemDescriptionField] = updatedBugBashItem.description;            
        fieldValues[teamFieldValue.field.referenceName] = teamFieldValue.defaultValue;        

        if (fieldValues["System.Tags-Add"]) {
            fieldValues["System.Tags"] = fieldValues["System.Tags-Add"];
        }            

        delete fieldValues["System.Tags-Add"];
        delete fieldValues["System.Tags-Remove"];

        try {
            // create work item
            savedWorkItem = await WorkItemActions.createWorkItem(bugBash.originalModel.workItemType, fieldValues);
        }
        catch (e) {
            BugBashItemActionsHub.UpdateBugBashItem.invoke({bugBashId: updatedBugBashItem.bugBashId, bugBashItem: updatedBugBashItem});
            throw e.message;
        }
        
        // associate work item with bug bash item
        updatedBugBashItem.workItemId = savedWorkItem.id;
        updatedBugBashItem.rejected = false;
        updatedBugBashItem.rejectedBy = "";
        updatedBugBashItem.rejectReason = "";
        updatedBugBashItem = await ExtensionDataManager.updateDocument(getBugBashCollectionKey(updatedBugBashItem.bugBashId), updatedBugBashItem, false);

        addExtraFieldsToWorkitem(savedWorkItem.id, updatedBugBashItem);

        return updatedBugBashItem;
    }

    function addExtraFieldsToWorkitem(workItemId: number, bugBashItem: IBugBashItem) {
        let fieldValues: IDictionaryStringTo<string> = {};
        const bugBash = StoresHub.bugBashStore.getItem(bugBashItem.bugBashId);

        fieldValues["System.History"] = getAcceptedItemComment(bugBash, bugBashItem);

        WorkItemActions.updateWorkItem(workItemId, fieldValues);
    }

    function getAcceptedItemComment(bugBash: BugBash, bugBashItem: IBugBashItem): string {
        const pageContext = Context.getPageContext();
        const navigation = pageContext.navigation;
        const webContext = VSS.getWebContext();
        const bugBashUrl = `${webContext.collection.uri}/${webContext.project.name}/_${navigation.currentController}/${navigation.currentAction}/${navigation.currentParameters}#_a=${UrlActions.ACTION_RESULTS}&id=${bugBash.id}`;
        
        const entity = parseUniquefiedIdentityName(bugBashItem.createdBy);

        let commentToSave = `
            Created from <a href='${bugBashUrl}' target='_blank'>${bugBash.originalModel.title}</a> bug bash on behalf of <a href='mailto:${entity.uniqueName || entity.displayName || ""}' data-vss-mention='version:1.0'>@${entity.displayName}</a>
        `;

        let discussionComments = StoresHub.bugBashItemCommentStore.getItem(bugBashItem.id);

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