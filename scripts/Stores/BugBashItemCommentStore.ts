import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");
import { WorkItemTemplate } from "TFS/WorkItemTracking/Contracts";
import * as WitClient from "TFS/WorkItemTracking/RestClient";

import { ExtensionDataManager } from "VSTS_Extension/Utilities/ExtensionDataManager";
import { BaseStore } from "VSTS_Extension/Stores/BaseStore";

import { ICommentDocument } from "../Models";
import { getBugBashItemCollectionKey } from "../Helpers";

export class BugBashItemCommentStore extends BaseStore<ICommentDocument[], ICommentDocument, string> {
    private _dataLoadedMap: IDictionaryStringTo<boolean>;

    constructor() {
        super();
        this.items = [];
        this._dataLoadedMap = {};
    }

    protected getItemByKey(id: string): ICommentDocument {
         return Utils_Array.first(this.items, (item: ICommentDocument) => Utils_String.equals(item.id, id, true));
    }

    protected async initializeItems(): Promise<void> {
        return;
    }

    public getKey(): string {
        return "BugBashItemCommentStore";
    }    

    public getItems(itemId: string): ICommentDocument[] {
        return this.items.filter(i => Utils_String.equals(i.bugBashItemDocumentId, itemId, true));
    }

    public isDataLoaded(itemId: string): boolean {
        return Boolean(this._dataLoadedMap[itemId.toLowerCase()]);
    }

    public async refreshItems(itemId: string) {
        const models = await ExtensionDataManager.readDocuments<ICommentDocument>(getBugBashItemCollectionKey(itemId), false);
        this.items = this.items.filter(i => !Utils_String.equals(i.bugBashItemDocumentId, itemId, true));
        this.items = this.items.concat(models);
        this._dataLoadedMap[itemId.toLowerCase()] = true;

        this.emitChanged();
    }

    public async addOrUpdateItem(item: ICommentDocument): Promise<ICommentDocument> {
        item.id = item.id || `${item.bugBashItemDocumentId}_${Date.now().toString()}`;
        const savedItem = await ExtensionDataManager.writeDocument(getBugBashItemCollectionKey(item.bugBashItemDocumentId), item, false);

        if (savedItem) {
            this._addItems(savedItem);
        }

        return savedItem;
    }

    private _addItems(items: ICommentDocument | ICommentDocument[]): void {
        if (!items) {
            return;
        }

        if (!this.items) {
            this.items = [];
        }

        if (Array.isArray(items)) {
            for (const item of items) {
                this._addItem(item);
            }
        }
        else {
            this._addItem(items);
        }

        this.emitChanged();
    }

    private _addItem(item: ICommentDocument): void {
        let existingItemIndex = Utils_Array.findIndex(this.items, (existingItem: ICommentDocument) => Utils_String.equals(item.id, existingItem.id, true));
        if (existingItemIndex != -1) {
            // Overwrite the item data
            this.items[existingItemIndex] = item;
        }
        else {
            this.items.push(item);
        }
    }
}