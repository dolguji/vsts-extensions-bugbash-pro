import { ExtensionDataManager } from "VSTS_Extension/Utilities/ExtensionDataManager";

import { WorkItem } from "TFS/WorkItemTracking/Contracts";

import { IBugBashItem, IAcceptedItemViewModel } from "./Interfaces";
import { createWorkItem, BugBashItemHelpers } from "./Helpers";
import { StoresHub } from "./Stores/StoresHub";

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
}

export class BugBashItemManager {
    public static async beginGetItems(bugBashId: string): Promise<IBugBashItem[]> {
        const models = await ExtensionDataManager.readDocuments<IBugBashItem>(getBugBashCollectionKey(bugBashId), false);
        for(let model of models) {
            translateDates(model);
            model.teamId = model.teamId || VSS.getWebContext().team.id;
        }
        return models;
    }

    public static async beginGetItem(itemId: string, bugBashId: string): Promise<IBugBashItem> {
        let model = await ExtensionDataManager.readDocument<IBugBashItem>(getBugBashCollectionKey(bugBashId), itemId, null, false);
        if (model) {
            translateDates(model);
            model.teamId = model.teamId || VSS.getWebContext().team.id;
            return model;
        }
        return null;
    }    

    public static async deleteItems(models: IBugBashItem[]): Promise<void> {
        await Promise.all(models.map(async (model) => {
            try {
                return await ExtensionDataManager.deleteDocument(getBugBashCollectionKey(model.bugBashId), model.id, false);
            }
            catch (e) {
                return null;
            }
        }));
    }

    public static async beginSave(model: IBugBashItem): Promise<IBugBashItem> {
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
            updatedItem = await this.beginSave(model);
        }
        catch (e) {
            throw "This item has been modified by some one else. Please refresh the item to get the latest version and try updating it again.";
        }

        // get accept template
        const bugBash = StoresHub.bugBashStore.getItem(model.bugBashId);
        const templateExists = await StoresHub.workItemTemplateItemStore.ensureTemplateItem(bugBash.acceptTemplate.templateId, bugBash.acceptTemplate.team);
        const teamFieldValueExists = await StoresHub.teamFieldStore.ensureItem(model.teamId);

        if (templateExists) {
            const template = StoresHub.workItemTemplateItemStore.getItem(bugBash.acceptTemplate.templateId);
            let fieldValues = {...template.fields};
            fieldValues["System.Title"] = model.title;            
            fieldValues[bugBash.itemDescriptionField] = model.description;

            if (teamFieldValueExists) {
                const teamFieldValue = StoresHub.teamFieldStore.getItem(model.teamId);
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
                    workItem: null
                };
            }
            
            // associate work item with bug bash item
            updatedItem.workItemId = savedWorkItem.id;
            updatedItem.rejected = false;
            updatedItem.rejectedBy = "";
            updatedItem.rejectReason = "";
            updatedItem = await this.beginSave(updatedItem);

            return {
                model: updatedItem,
                workItem: savedWorkItem
            };
        }
        else {
            return {
                model: updatedItem,
                workItem: null
            }
        }
    }
}