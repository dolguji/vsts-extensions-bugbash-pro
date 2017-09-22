import { ExtensionDataManager } from "MB/Utils/ExtensionDataManager";
import { parseUniquefiedIdentityName } from "MB/Components/IdentityView";
import { WorkItemTemplateItemActions } from "MB/Flux/Actions/WorkItemTemplateItemActions";
import { TeamFieldActions } from "MB/Flux/Actions/TeamFieldActions";
import { WorkItemActions } from "MB/Flux/Actions/WorkItemActions";
import { DateUtils } from "MB/Utils/Date";
import { StringUtils } from "MB/Utils/String";

import { WorkItem, WorkItemTemplate } from "TFS/WorkItemTracking/Contracts";

import { UrlActions, BugBashFieldNames, ErrorKeys } from "../Constants";
import { BugBashItemActionsHub, BugBashErrorMessageActionsHub, BugBashClientActionsHub } from "./ActionsHub";
import { StoresHub } from "../Stores/StoresHub";
import { IBugBashItem, IBugBashItemComment } from "../Interfaces";
import { BugBash } from "../ViewModels/BugBash";
import { BugBashItemCommentActions } from "../Actions/BugBashItemCommentActions";
import { getBugBashUrl } from "../Helpers";

export module BugBashItemActions {
    export function fireStoreChange() {
        BugBashItemActionsHub.FireStoreChange.invoke(null);
    }

    export async function initializeItems(bugBashId: string) {
        if (StoresHub.bugBashItemStore.isLoaded(bugBashId)) {
            BugBashItemActionsHub.InitializeBugBashItems.invoke(null);
        }
        else if (!StoresHub.bugBashItemStore.isLoading(bugBashId)) {
            StoresHub.bugBashItemStore.setLoading(true, bugBashId);

            const bugBashItemModels = await ExtensionDataManager.readDocuments<IBugBashItem>(getBugBashCollectionKey(bugBashId), false);
            for(let bugBashItemModel of bugBashItemModels) {
                preProcessModel(bugBashItemModel);
            }
            
            BugBashItemActionsHub.InitializeBugBashItems.invoke({
                bugBashId: bugBashId,
                bugBashItemModels: bugBashItemModels
            });
            StoresHub.bugBashItemStore.setLoading(false, bugBashId);
        }
    } 

    export async function refreshItems(bugBashId: string) {
        if (!StoresHub.bugBashItemStore.isLoading(bugBashId)) {
            StoresHub.bugBashItemStore.setLoading(true, bugBashId);

            const bugBashItemModels = await ExtensionDataManager.readDocuments<IBugBashItem>(getBugBashCollectionKey(bugBashId), false);
            for(let bugBashItemModel of bugBashItemModels) {
                preProcessModel(bugBashItemModel);
            }
            
            BugBashItemActionsHub.RefreshBugBashItems.invoke({
                bugBashId: bugBashId,
                bugBashItemModels: bugBashItemModels
            });
            StoresHub.bugBashItemStore.setLoading(false, bugBashId);

            BugBashErrorMessageActionsHub.DismissErrorMessage.invoke(ErrorKeys.BugBashItemError);

            // clear all comments
            BugBashItemCommentActions.clearComments();
        }
    } 

    export async function refreshItem(bugBashId: string, bugBashItemId: string) {
        if (!StoresHub.bugBashItemStore.isLoading(bugBashId) && !StoresHub.bugBashItemStore.isLoading(bugBashItemId)) {
            StoresHub.bugBashItemStore.setLoading(true, bugBashItemId);
            const bugBashItemModel = await ExtensionDataManager.readDocument<IBugBashItem>(getBugBashCollectionKey(bugBashId), bugBashItemId, null, false);

            if (bugBashItemModel) {
                preProcessModel(bugBashItemModel);
                
                BugBashItemActionsHub.RefreshBugBashItem.invoke({
                    bugBashId: bugBashId,
                    bugBashItemModel: bugBashItemModel
                });
                StoresHub.bugBashItemStore.setLoading(false, bugBashItemId);
                BugBashErrorMessageActionsHub.DismissErrorMessage.invoke(ErrorKeys.BugBashItemError);

                // refresh comments for this bug bash item
                BugBashItemCommentActions.refreshComments(bugBashItemId);
            }
            else {
                StoresHub.bugBashItemStore.setLoading(false, bugBashItemId);
                BugBashErrorMessageActionsHub.PushErrorMessage.invoke({
                    errorMessage: "This instance of bug bash item does not exist.",
                    errorKey: ErrorKeys.BugBashItemError
                });
            }
        }
    }

    export async function updateBugBashItem(bugBashId: string, bugBashItemModel: IBugBashItem, newComment?: string) {
        if (!StoresHub.bugBashItemStore.isLoading(bugBashId) && !StoresHub.bugBashItemStore.isLoading(bugBashItemModel.id)) {
            StoresHub.bugBashItemStore.setLoading(true, bugBashItemModel.id);

            try {
                let updatedBugBashItemModel = await ExtensionDataManager.updateDocument(getBugBashCollectionKey(bugBashId), bugBashItemModel, false);
                preProcessModel(updatedBugBashItemModel);
                
                BugBashItemActionsHub.UpdateBugBashItem.invoke({
                    bugBashId: bugBashId,
                    bugBashItemModel: updatedBugBashItemModel
                });
                StoresHub.bugBashItemStore.setLoading(false, bugBashItemModel.id);
                BugBashErrorMessageActionsHub.DismissErrorMessage.invoke(ErrorKeys.BugBashItemError);

                if (newComment != null && newComment.trim() !== "") {
                    BugBashItemCommentActions.createComment(updatedBugBashItemModel.id, newComment);
                }

                BugBashClientActionsHub.SelectedBugBashItemChanged.invoke(updatedBugBashItemModel.id);
            }
            catch (e) {
                StoresHub.bugBashItemStore.setLoading(false, bugBashItemModel.id);
                BugBashErrorMessageActionsHub.PushErrorMessage.invoke({
                    errorMessage: "This bug bash item has been modified by some one else. Please refresh the item to get the latest version and try updating it again.",
                    errorKey: ErrorKeys.BugBashItemError
                });
            }
        }
    }

    export async function createBugBashItem(bugBashId: string, bugBashItemModel: IBugBashItem, newComment?: string) {
        if (!StoresHub.bugBashItemStore.isLoading(bugBashId)) {
            try {
                let cloneBugBashItemModel = {...bugBashItemModel};
                cloneBugBashItemModel.bugBashId = bugBashId;
                cloneBugBashItemModel.createdBy = `${VSS.getWebContext().user.name} <${VSS.getWebContext().user.uniqueName}>`;
                cloneBugBashItemModel.createdDate = new Date(Date.now());
                const bugBash = StoresHub.bugBashStore.getItem(bugBashId);
                let createdBugBashItemModel: IBugBashItem;

                if (bugBash.isAutoAccept) {
                    createdBugBashItemModel = await acceptItem(cloneBugBashItemModel, newComment);
                }
                else {
                    createdBugBashItemModel = await ExtensionDataManager.createDocument(getBugBashCollectionKey(bugBashId), cloneBugBashItemModel, false);
                }                
                preProcessModel(createdBugBashItemModel);
                
                BugBashItemActionsHub.CreateBugBashItem.invoke({
                    bugBashId: bugBashId,
                    bugBashItemModel: createdBugBashItemModel
                });

                BugBashClientActionsHub.SelectedBugBashItemChanged.invoke(createdBugBashItemModel.id);
                BugBashErrorMessageActionsHub.DismissErrorMessage.invoke(ErrorKeys.BugBashItemError);

                if (newComment != null && newComment.trim() !== "") {
                    BugBashItemCommentActions.createComment(createdBugBashItemModel.id, newComment);
                }
            }
            catch (e) {
                BugBashErrorMessageActionsHub.PushErrorMessage.invoke({
                    errorMessage: e.message,
                    errorKey: ErrorKeys.BugBashItemError
                });
            }
        }
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

            BugBashItemActionsHub.DeleteBugBashItem.invoke({
                bugBashId: bugBashId, 
                bugBashItemId: bugBashItemId
            });
            StoresHub.bugBashItemStore.setLoading(false, bugBashItemId);
        }
    }

    export async function acceptBugBashItem(bugBashId: string, bugBashItemModel: IBugBashItem) {
        if (!StoresHub.bugBashItemStore.isLoading(bugBashId) && !StoresHub.bugBashItemStore.isLoading(bugBashItemModel.id)) {
            StoresHub.bugBashItemStore.setLoading(true, bugBashItemModel.id);

            try {
                let acceptedBugBashItemModel = await acceptItem(bugBashItemModel);
                preProcessModel(acceptedBugBashItemModel);

                BugBashItemActionsHub.AcceptBugBashItem.invoke({
                    bugBashId: bugBashId,
                    bugBashItemModel: acceptedBugBashItemModel
                });
                StoresHub.bugBashItemStore.setLoading(false, bugBashItemModel.id);

                BugBashClientActionsHub.SelectedBugBashItemChanged.invoke(acceptedBugBashItemModel.id);
                BugBashErrorMessageActionsHub.DismissErrorMessage.invoke(ErrorKeys.BugBashItemError);                
            }
            catch (e) {
                StoresHub.bugBashItemStore.setLoading(false, bugBashItemModel.id);
                BugBashErrorMessageActionsHub.PushErrorMessage.invoke({
                    errorMessage: e,
                    errorKey: ErrorKeys.BugBashItemError
                });
            }
        }
    }    

    function getBugBashCollectionKey(bugBashId: string): string {
        return StringUtils.isGuid(bugBashId) ? `Items_${bugBashId}` : `BugBashCollection_${bugBashId}`;
    }

    function preProcessModel(bugBashItem: IBugBashItem) {
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

    async function acceptItem(bugBashItemModel: IBugBashItem, newComment?: string): Promise<IBugBashItem> {
        let updatedBugBashItemModel: IBugBashItem;
        let savedWorkItem: WorkItem;
        const bugBash = StoresHub.bugBashStore.getItem(bugBashItemModel.bugBashId);
        let acceptTemplate: WorkItemTemplate;

        // read bug bash wit template
        const acceptTemplateId = bugBash.getFieldValue<string>(BugBashFieldNames.AcceptTemplateId, true);
        const acceptTemplateTeam = bugBash.getFieldValue<string>(BugBashFieldNames.AcceptTemplateTeam, true);

        if (acceptTemplateTeam && acceptTemplateId) {
            try {
                await WorkItemTemplateItemActions.initializeWorkItemTemplateItem(acceptTemplateTeam, acceptTemplateId);
                acceptTemplate = StoresHub.workItemTemplateItemStore.getItem(acceptTemplateId);
            }
            catch (e) {
                throw `Bug bash template '${acceptTemplateId}' does not exist in team '${acceptTemplateTeam}'`;
            }
        }

        try {
            await TeamFieldActions.initializeTeamFields(bugBashItemModel.teamId);
        }
        catch (e) {
            throw `Cannot read team field value for team: ${bugBashItemModel.teamId}.`;
        }

        if (bugBashItemModel.__etag) {
            try {
                // first do a empty save to check if its the latest version of the item
                updatedBugBashItemModel = await ExtensionDataManager.updateDocument(getBugBashCollectionKey(bugBashItemModel.bugBashId), bugBashItemModel, false);
            }
            catch (e) {
                throw "This item has been modified by some one else. Please refresh the item to get the latest version and try updating it again.";
            }
        }
        else {
            updatedBugBashItemModel = {...bugBashItemModel}; // For auto accept scenario
        }

        const itemDescriptionField = bugBash.getFieldValue<string>(BugBashFieldNames.ItemDescriptionField, true);
        const workItemType = bugBash.getFieldValue<string>(BugBashFieldNames.WorkItemType, true);
        const teamFieldValue = StoresHub.teamFieldStore.getItem(updatedBugBashItemModel.teamId);

        let fieldValues = acceptTemplate ? {...acceptTemplate.fields} : {};
        fieldValues["System.Title"] = updatedBugBashItemModel.title;
        fieldValues[itemDescriptionField] = updatedBugBashItemModel.description;            
        fieldValues[teamFieldValue.field.referenceName] = teamFieldValue.defaultValue;        

        if (fieldValues["System.Tags-Add"]) {
            fieldValues["System.Tags"] = fieldValues["System.Tags-Add"];
        }            

        delete fieldValues["System.Tags-Add"];
        delete fieldValues["System.Tags-Remove"];

        try {
            // create work item
            savedWorkItem = await WorkItemActions.createWorkItem(workItemType, fieldValues);
        }
        catch (e) {
            BugBashItemActionsHub.UpdateBugBashItem.invoke({bugBashId: updatedBugBashItemModel.bugBashId, bugBashItemModel: updatedBugBashItemModel});
            throw e;
        }
        
        // associate work item with bug bash item
        updatedBugBashItemModel.workItemId = savedWorkItem.id;
        updatedBugBashItemModel.rejected = false;
        updatedBugBashItemModel.rejectedBy = "";
        updatedBugBashItemModel.rejectReason = "";

        if (updatedBugBashItemModel.__etag) {
            updatedBugBashItemModel = await ExtensionDataManager.updateDocument(getBugBashCollectionKey(updatedBugBashItemModel.bugBashId), updatedBugBashItemModel, false);
        }
        else {
            updatedBugBashItemModel = await ExtensionDataManager.createDocument(getBugBashCollectionKey(updatedBugBashItemModel.bugBashId), updatedBugBashItemModel, false);
        }

        addExtraFieldsToWorkitem(savedWorkItem.id, updatedBugBashItemModel, newComment);

        return updatedBugBashItemModel;
    }

    function addExtraFieldsToWorkitem(workItemId: number, bugBashItemModel: IBugBashItem, newComment?: string) {
        let fieldValues: IDictionaryStringTo<string> = {};
        const bugBash = StoresHub.bugBashStore.getItem(bugBashItemModel.bugBashId);

        fieldValues["System.History"] = getAcceptedItemComment(bugBash, bugBashItemModel, newComment);

        WorkItemActions.updateWorkItem(workItemId, fieldValues);
    }

    function getAcceptedItemComment(bugBash: BugBash, bugBashItemModel: IBugBashItem, newComment?: string): string {
        const entity = parseUniquefiedIdentityName(bugBashItemModel.createdBy);
        const bugBashTitle = bugBash.getFieldValue<string>(BugBashFieldNames.Title, true);

        let commentToSave = `
            Created from <a href='${getBugBashUrl(bugBash.id, UrlActions.ACTION_RESULTS)}' target='_blank'>${bugBashTitle}</a> bug bash on behalf of <a href='mailto:${entity.uniqueName || entity.displayName || ""}' data-vss-mention='version:1.0'>@${entity.displayName}</a>
        `;

        let discussionComments = newComment ? [{
            content: newComment,
            createdDate: new Date(Date.now()),
            createdBy: `${VSS.getWebContext().user.name} <${VSS.getWebContext().user.uniqueName}>`
        }] : StoresHub.bugBashItemCommentStore.getItem(bugBashItemModel.id);

        if (discussionComments && discussionComments.length > 0) {
            discussionComments = discussionComments.slice();
            discussionComments = discussionComments.sort((c1: IBugBashItemComment, c2: IBugBashItemComment) => DateUtils.defaultComparer(c1.createdDate, c2.createdDate));

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