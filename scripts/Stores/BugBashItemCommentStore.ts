import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");

import { ExtensionDataManager } from "VSTS_Extension/Utilities/ExtensionDataManager";
import { BaseStore } from "VSTS_Extension/Stores/BaseStore";

import { IBugBashItemComment } from "../Models";
import { getBugBashItemCollectionKey } from "../Helpers";

export class BugBashItemCommentStore extends BaseStore<IBugBashItemComment[], IBugBashItemComment, string> {
    private _dataLoadedMap: IDictionaryStringTo<boolean>;

    constructor() {
        super();
        this.items = [];
        this._dataLoadedMap = {};
    }

    protected getItemByKey(id: string): IBugBashItemComment {
         return Utils_Array.first(this.items, (item: IBugBashItemComment) => Utils_String.equals(item.id, id, true));
    }

    protected async initializeItems(): Promise<void> {
        return;
    }

    public getKey(): string {
        return "BugBashItemCommentStore";
    }    

    public getItems(itemId: string): IBugBashItemComment[] {
        return this.items.filter(i => Utils_String.equals(i.bugBashItemId, itemId, true));
    }

    public isDataLoaded(itemId: string): boolean {
        return Boolean(this._dataLoadedMap[itemId.toLowerCase()]);
    }

    public clear(): void {
        this._dataLoadedMap = {};
        this.items = [];
        this.emitChanged();
    }

    public async refreshItems(itemId: string) {
        delete this._dataLoadedMap[itemId.toLowerCase()];
        this._removeItems(this.items.filter(i => Utils_String.equals(i.bugBashItemId, itemId, true)));

        const models = await ExtensionDataManager.readDocuments<IBugBashItemComment>(getBugBashItemCollectionKey(itemId), false);
        for(let model of models) {
            this._translateDates(model);
        }

        this._dataLoadedMap[itemId.toLowerCase()] = true;
        this._addItems(models);
    }

    public async createItem(item: IBugBashItemComment): Promise<IBugBashItemComment> {
        item.id = item.id || `${item.bugBashItemId}_${Date.now().toString()}`;
        item.addedBy = item.addedBy || `${VSS.getWebContext().user.name} <${VSS.getWebContext().user.uniqueName}>`;
        item.addedDate = item.addedDate || new Date(Date.now());

        const savedItem = await ExtensionDataManager.createDocument(getBugBashItemCollectionKey(item.bugBashItemId), item, false);

        if (savedItem) {
            this._translateDates(savedItem);
            this._addItems(savedItem);
        }

        return savedItem;
    }

    private _addItems(items: IBugBashItemComment | IBugBashItemComment[]): void {
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

    private _addItem(item: IBugBashItemComment): void {
        let existingItemIndex = Utils_Array.findIndex(this.items, (existingItem: IBugBashItemComment) => Utils_String.equals(item.id, existingItem.id, true));
        if (existingItemIndex != -1) {
            // Overwrite the item data
            this.items[existingItemIndex] = item;
        }
        else {
            this.items.push(item);
        }
    }

    private _removeItems(items: IBugBashItemComment | IBugBashItemComment[]): void {
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

    private _removeItem(item: IBugBashItemComment): void {
        const existingItemIndex = Utils_Array.findIndex(this.items, (existingItem: IBugBashItemComment) => Utils_String.equals(item.id, existingItem.id, true));

        if (existingItemIndex != -1) {
            this.items.splice(existingItemIndex, 1);
        }
    }

    private _translateDates(item: IBugBashItemComment) {
        if (typeof item.addedDate === "string") {
            if ((item.addedDate as string).trim() === "") {
                item.addedDate = undefined;
            }
            else {
                item.addedDate = new Date(item.addedDate);
            }
        }
    }
}