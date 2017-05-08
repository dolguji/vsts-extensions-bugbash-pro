import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");
import { WorkItemTemplate } from "TFS/WorkItemTracking/Contracts";
import * as WitClient from "TFS/WorkItemTracking/RestClient";

import { ExtensionDataManager } from "VSTS_Extension/Utilities/ExtensionDataManager";
import { BaseStore } from "VSTS_Extension/Stores/BaseStore";

import { IBugBashItemDocument } from "../Models";
import { getBugBashCollectionKey } from "../Helpers";

export class BugBashItemStore extends BaseStore<IBugBashItemDocument[], IBugBashItemDocument, string> {
    constructor() {
        super();
        this.items = [];    
    }

    protected getItemByKey(id: string): IBugBashItemDocument {
         return Utils_Array.first(this.items, (item: IBugBashItemDocument) => Utils_String.equals(item.id, id, true));
    }

    protected async initializeItems(): Promise<void> {
        return;
    }

    public getKey(): string {
        return "BugBashItemStore";
    }    

    public getItems(bugBashId: string): IBugBashItemDocument[] {
        return this.items.filter(i => Utils_String.equals(i.bugBashId, bugBashId, true));
    }

    public async refreshItems(bugBashId: string) {
        const models = await ExtensionDataManager.readDocuments<IBugBashItemDocument>(getBugBashCollectionKey(bugBashId), false);
        this.items = this.items.filter(i => !Utils_String.equals(i.bugBashId, bugBashId, true));
        this.items = this.items.concat(models);
        this.emitChanged();
    }

    public async addOrUpdateItem(item: IBugBashItemDocument): Promise<IBugBashItemDocument> {
        item.id = item.id || Date.now().toString();    
        const savedItem = await ExtensionDataManager.writeDocument(getBugBashCollectionKey(item.bugBashId), item, false);

        if (savedItem) {
            this._addItems(savedItem);
        }

        return savedItem;
    }

    public async deleteItem(item: IBugBashItemDocument): Promise<void> {
        this._removeItems(item);
        await ExtensionDataManager.deleteDocument(getBugBashCollectionKey(item.bugBashId), item.id, false);        
    }

    public async ensureItem(bugBashId: string, id: string): Promise<boolean> {
        if (!this.itemExists(id)) {
            try {
                const model = await ExtensionDataManager.readDocument<IBugBashItemDocument>(getBugBashCollectionKey(bugBashId), id, null, false);
                if (model) {
                    this._addItems(model);
                    return true;
                }
            }
            catch (e) {
                return false;
            }

            return false;
        }
        else {
            return true;
        }
    }

    private _addItems(items: IBugBashItemDocument | IBugBashItemDocument[]): void {
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

    private _removeItems(items: IBugBashItemDocument | IBugBashItemDocument[]): void {
        if (!items || !this.isLoaded()) {
            return;
        }

        if (Array.isArray(items)) {
            for (const item of items) {
                this._removeItem(item);
            }
        }
        else {
            this._removeItem(items);
        }

        this.emitChanged();
    }    

    private _addItem(item: IBugBashItemDocument): void {
        let existingItemIndex = Utils_Array.findIndex(this.items, (existingItem: IBugBashItemDocument) => Utils_String.equals(item.id, existingItem.id, true));
        if (existingItemIndex != -1) {
            // Overwrite the item data
            this.items[existingItemIndex] = item;
        }
        else {
            this.items.push(item);
        }
    }

    private _removeItem(item: IBugBashItemDocument): void {
        const existingItemIndex = Utils_Array.findIndex(this.items, (existingItem: IBugBashItemDocument) => Utils_String.equals(item.id, existingItem.id, true));

        if (existingItemIndex != -1) {
            this.items.splice(existingItemIndex, 1);
        }
    }
}