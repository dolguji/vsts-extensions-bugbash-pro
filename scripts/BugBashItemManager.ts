import { IBugBashItem, IBugBashItemModel } from "./Models";
import { ExtensionDataManager } from "VSTS_Extension/Utilities/ExtensionDataManager";
import Utils_Array = require("VSS/Utils/Array");
import Utils_String = require("VSS/Utils/String");

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

    if (typeof model.acceptedDate === "string") {
        if ((model.acceptedDate as string).trim() === "") {
            model.acceptedDate = undefined;
        }
        else {
            model.acceptedDate = new Date(model.acceptedDate);
        }
    }

    for (let i = 0; i < model.comments.length; i++) {
        if (typeof model.comments[i].addedDate === "string") {
            if ((model.comments[i].addedDate as any).trim() === "") {
                model.comments[i].addedDate = undefined;
            }
            else {
                model.comments[i].addedDate = new Date(model.comments[i].addedDate as any);
            }
        }
    }
}

export class BugBashItemManager {
    public static getNewItemModel(bugBashId: string): IBugBashItemModel {        
        return {
            model: this.getNewItem(bugBashId),
            originalModel: this.getNewItem(bugBashId),
            newComment: ""
        }
    }

    public static getNewItem(bugBashId: string): IBugBashItem {
        return {
            id: "",
            bugBashId: bugBashId,
            __etag: 0,
            title: "",
            comments: [],
            description: "",
            workItemId: 0,
            createdDate: null,
            createdBy: "",
            acceptedDate: null,
            acceptedBy: ""
        };
    }

    public static deepCopy(model: IBugBashItem): IBugBashItem {
        return {
            id: model.id,
            bugBashId: model.bugBashId,
            __etag: model.__etag,
            title: model.title,
            comments: model.comments.slice(),
            description: model.description,
            workItemId: model.workItemId,
            createdDate: model.createdDate,
            createdBy: model.createdBy,
            acceptedDate: model.acceptedDate,
            acceptedBy: model.acceptedBy
        }    
    }

    public static getItemModel(model: IBugBashItem): IBugBashItemModel {
        return {
            model: this.deepCopy(model),
            originalModel: this.deepCopy(model),
            newComment: ""
        }
    }

    public static async beginGetItems(bugBashId: string): Promise<IBugBashItem[]> {
        const models = await ExtensionDataManager.readDocuments<IBugBashItem>(getBugBashCollectionKey(bugBashId), false);
        for(let model of models) {
            translateDates(model);
        }
        return models;
    }

    public static async beginGetItem(itemId: string, bugBashId: string): Promise<IBugBashItem> {
        let model = await ExtensionDataManager.readDocument<IBugBashItem>(getBugBashCollectionKey(bugBashId), itemId, null, false);
        if (model) {
            translateDates(model);
            return model;
        }
        return null;
    }    

    public static async deleteItems(items: IBugBashItem[]): Promise<void> {
        await Promise.all(items.map(async (item) => {
            try {
                return await ExtensionDataManager.deleteDocument(getBugBashCollectionKey(item.bugBashId), item.id, false);
            }
            catch (e) {
                return null;
            }
        }));
    }

    public static isNew(model: IBugBashItem): boolean {
        return !model.id;
    }

    public static isDirty(itemModel: IBugBashItemModel): boolean {        
        return !Utils_String.equals(itemModel.model.title, itemModel.originalModel.title)
            || !Utils_String.equals(itemModel.model.description, itemModel.originalModel.description)
            || itemModel.newComment !== "";
    }

    public static isValid(model: IBugBashItem): boolean {
        return model.title.trim().length > 0 && model.title.trim().length <= 256;
    }

    public static isAccepted(model: IBugBashItem): boolean {
        return model.workItemId != null && model.workItemId > 0;
    }

    public static async beginSave(model: IBugBashItem, newComment?: string): Promise<IBugBashItem> {
        const isNew = this.isNew(model);
        let cloneModel = this.deepCopy(model);
        let updatedModel: IBugBashItem;

        cloneModel.id = model.id || `${model.bugBashId}_${Date.now().toString()}`;
        cloneModel.createdBy = model.createdBy || `${VSS.getWebContext().user.name} <${VSS.getWebContext().user.uniqueName}>`;
        cloneModel.createdDate = model.createdDate || new Date(Date.now());
        cloneModel.comments = model.comments || [];

        if (newComment && newComment.trim() !== "") {
            cloneModel.comments = cloneModel.comments.concat([{
                text: newComment,
                addedBy: `${VSS.getWebContext().user.name} <${VSS.getWebContext().user.uniqueName}>`,
                addedDate: new Date(Date.now())
            }]);
        }

        if (isNew) {
            updatedModel = await ExtensionDataManager.createDocument(getBugBashCollectionKey(cloneModel.bugBashId), cloneModel, false);
        }
        else {
            updatedModel = await ExtensionDataManager.updateDocument(getBugBashCollectionKey(cloneModel.bugBashId), cloneModel, false);
        }

        translateDates(updatedModel);

        return updatedModel;
    }
}