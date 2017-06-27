import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");

import { BaseStore } from "VSTS_Extension/Stores/BaseStore";
import { ExtensionDataManager } from "VSTS_Extension/Utilities/ExtensionDataManager";

import { IBugBashItemComment } from "../Interfaces";

function getBugBashItemCollectionKey(itemId: string): string {
    return `BugBashItemCollection_${itemId}`;
}

function translateDates(model: IBugBashItemComment) {
    if (typeof model.createdDate === "string") {
        if ((model.createdDate as string).trim() === "") {
            model.createdDate = undefined;
        }
        else {
            model.createdDate = new Date(model.createdDate);
        }
    }
}

export class BugBashItemCommentStore extends BaseStore<IDictionaryStringTo<IBugBashItemComment[]>, IBugBashItemComment[], string> {
    constructor() {
        super();
        this.items = {};    
    }

    protected getItemByKey(bugBashItemId: string): IBugBashItemComment[] {
         return this.items[bugBashItemId.toLowerCase()];
    }

    protected async initializeItems(): Promise<void> {
        return;   
    }

    public getKey(): string {
        return "BugBashItemCommentStore";
    }

    public async ensureComments(bugBashItemId: string): Promise<void> {
        if (!this.itemExists(bugBashItemId)) {
            const models = await ExtensionDataManager.readDocuments<IBugBashItemComment>(getBugBashItemCollectionKey(bugBashItemId), false);
            for(let model of models) {
                translateDates(model);
            }
            this.items[bugBashItemId.toLowerCase()] = models;
            this.emitChanged();
        }
        else {
            this.emitChanged();
        }
    }

    public refreshComments(bugBashItemId: string) {
        delete this.items[bugBashItemId.toLowerCase()];
        this.emitChanged();
        this.ensureComments(bugBashItemId);
    }

    public async addCommentItem(bugBashItemId: string, newComment: string): Promise<void> {
        let commentModel: IBugBashItemComment = {
            id: `${bugBashItemId}_${Date.now().toString()}`,
            __etag: 0,
            createdBy: `${VSS.getWebContext().user.name} <${VSS.getWebContext().user.uniqueName}>`,
            createdDate: new Date(Date.now()),
            content: newComment
        };

        const savedComment = await ExtensionDataManager.createDocument<IBugBashItemComment>(getBugBashItemCollectionKey(bugBashItemId), commentModel, false);
        translateDates(savedComment);

        let comments = this.getItemByKey(bugBashItemId) || [];
        comments.push(savedComment);

        this.items[bugBashItemId.toLowerCase()] = comments;

        this.emitChanged();
    }
}