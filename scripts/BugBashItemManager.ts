import { ExtensionDataManager } from "VSTS_Extension/Utilities/ExtensionDataManager";
import { parseUniquefiedIdentityName } from "VSTS_Extension/Components/WorkItemControls/IdentityView";

import { WorkItem } from "TFS/WorkItemTracking/Contracts";
import Context = require("VSS/Context");
import Utils_Date = require("VSS/Utils/Date");

import { UrlActions } from "./Constants";
import { IBugBash, IBugBashItem, IAcceptedItemViewModel, IBugBashItemComment } from "./Interfaces";
import { createWorkItem, updateWorkItem, BugBashItemHelpers } from "./Helpers";
import { StoresHub } from "./Stores/StoresHub";


export class BugBashItemManager {
    public static async saveItem(model: IBugBashItem): Promise<IBugBashItem> {
        const isNew = BugBashItemHelpers.isNew(model);
        let cloneModel = BugBashItemHelpers.deepCopy(model);
        let updatedModel: IBugBashItem;

        cloneModel.id = model.id || `${model.bugBashId}_${Date.now().toString()}`;
        cloneModel.createdBy = model.createdBy || `${VSS.getWebContext().user.name} <${VSS.getWebContext().user.uniqueName}>`;
        cloneModel.createdDate = model.createdDate || new Date(Date.now());

        if (isNew) {
            updatedModel = await ExtensionDataManager.createDocument(getBugBashCollectionKey(cloneModel.bugBashId), cloneModel, false);
        }
        else {
            updatedModel = await ExtensionDataManager.updateDocument(getBugBashCollectionKey(cloneModel.bugBashId), cloneModel, false);
        }

        translateDates(updatedModel);

        return updatedModel;
    }

    public static async acceptItem(model: IBugBashItem): Promise<IAcceptedItemViewModel> {
        let updatedItem: IBugBashItem;
        let savedWorkItem: WorkItem;

        try {
            // first do a empty save to check if its the latest version of the item
            updatedItem = await this.saveItem(model);
        }
        catch (e) {
            throw "This item has been modified by some one else. Please refresh the item to get the latest version and try updating it again.";
        }

        // get accept template
        const bugBash = StoresHub.bugBashStore.getItem(updatedItem.bugBashId);
        const templateExists = await StoresHub.workItemTemplateItemStore.ensureTemplateItem(bugBash.acceptTemplate.templateId, bugBash.acceptTemplate.team);
        const teamFieldValueExists = await StoresHub.teamFieldStore.ensureItem(updatedItem.teamId);

        if (templateExists) {
            const template = StoresHub.workItemTemplateItemStore.getItem(bugBash.acceptTemplate.templateId);
            let fieldValues = {...template.fields};
            fieldValues["System.Title"] = updatedItem.title;
            fieldValues[bugBash.itemDescriptionField] = updatedItem.description;            

            if (teamFieldValueExists) {
                const teamFieldValue = StoresHub.teamFieldStore.getItem(updatedItem.teamId);
                fieldValues[teamFieldValue.field.referenceName] = teamFieldValue.defaultValue;
            }

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
                return {
                    model: updatedItem,
                    workItem: null,
                    error: e.message
                };
            }
            
            // associate work item with bug bash item
            updatedItem.workItemId = savedWorkItem.id;
            updatedItem.rejected = false;
            updatedItem.rejectedBy = "";
            updatedItem.rejectReason = "";
            updatedItem = await this.saveItem(updatedItem);

            this._addExtraFieldsToWorkitem(savedWorkItem.id, updatedItem);

            return {
                model: updatedItem,
                workItem: savedWorkItem
            };
        }
        else {
            return {
                model: updatedItem,
                workItem: null,
                error: `Bug bash template '${bugBash.acceptTemplate.templateId}' does not exist in team '${bugBash.acceptTemplate.team}'`
            }
        }
    }

    private static _addExtraFieldsToWorkitem(workItemId: number, model: IBugBashItem) {
        let fieldValues: IDictionaryStringTo<string> = {};
        const bugBash = StoresHub.bugBashStore.getItem(model.bugBashId);

        fieldValues["System.History"] = this._getAcceptedItemComment(bugBash, model);

        updateWorkItem(workItemId, fieldValues);
    }

    private static _getAcceptedItemComment(bugBash: IBugBash, model: IBugBashItem): string {
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