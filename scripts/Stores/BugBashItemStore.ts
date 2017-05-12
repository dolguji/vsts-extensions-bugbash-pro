import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");

import { ExtensionDataManager } from "VSTS_Extension/Utilities/ExtensionDataManager";
import { BaseStore } from "VSTS_Extension/Stores/BaseStore";

import { IBugBashItem } from "../Models";
import { getBugBashCollectionKey } from "../Helpers";

export class BugBashItemStore extends BaseStore<IBugBashItem[], IBugBashItem, string> {
    private _dataLoadedMap: IDictionaryStringTo<boolean>;

    constructor() {
        super();
        this.items = [];
        this._dataLoadedMap = {};
    }

    protected getItemByKey(id: string): IBugBashItem {
         return Utils_Array.first(this.items, (item: IBugBashItem) => Utils_String.equals(item.id, id, true));
    }

    protected async initializeItems(): Promise<void> {
        return;
    }

    public getKey(): string {
        return "BugBashItemStore";
    }    

    public getItems(bugBashId: string): IBugBashItem[] {
        return this.items.filter(i => Utils_String.equals(i.bugBashId, bugBashId, true));
    }

    public isDataLoaded(bugBashId: string): boolean {
        return Boolean(this._dataLoadedMap[bugBashId]);
    }

    public async refreshItems(bugBashId: string) {
        delete this._dataLoadedMap[bugBashId];
        this._removeItems(this.items.filter(i => Utils_String.equals(i.bugBashId, bugBashId, true)));

        const models = await ExtensionDataManager.readDocuments<IBugBashItem>(getBugBashCollectionKey(bugBashId), false);
        for(let model of models) {
            this._translateDates(model);
        }
        
        this._dataLoadedMap[bugBashId] = true;
        this._addItems(models);
    }

    public async refreshItem(item: IBugBashItem): Promise<IBugBashItem> {
        let refreshedItem = await ExtensionDataManager.readDocument<IBugBashItem>(getBugBashCollectionKey(item.bugBashId), item.id, null, false);
        if (refreshedItem) {
            this._addItems(refreshedItem);
        }
        else {
            this._removeItem(refreshedItem);
        }
        return refreshedItem;
    }

    public async createItem(item: IBugBashItem): Promise<IBugBashItem> {
        let model = {...item};
        model.id = model.id || `${model.bugBashId}_${Date.now().toString()}`;
        model.createdBy = model.createdBy || `${VSS.getWebContext().user.name} <${VSS.getWebContext().user.uniqueName}>`;
        model.createdDate = model.createdDate || new Date(Date.now());
        const savedItem = await ExtensionDataManager.createDocument(getBugBashCollectionKey(model.bugBashId), model, false);

        this._translateDates(savedItem);
        this._addItems(savedItem);

        return savedItem;
    }

    public async updateItem(item: IBugBashItem): Promise<IBugBashItem> {        
        const savedItem = await ExtensionDataManager.updateDocument(getBugBashCollectionKey(item.bugBashId), item, false);
        this._translateDates(savedItem);
        this._addItems(savedItem);
        return savedItem;
    }

    public async deleteItems(items: IBugBashItem[]): Promise<void> {
        this._removeItems(items);
        await Promise.all(items.map(async (item) => {
            try {
                return await ExtensionDataManager.deleteDocument(getBugBashCollectionKey(item.bugBashId), item.id, false);
            }
            catch (e) {
                return null;
            }
        }));        
    }

    private _addItems(items: IBugBashItem | IBugBashItem[]): void {
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

    private _removeItems(items: IBugBashItem | IBugBashItem[]): void {
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

    private _addItem(item: IBugBashItem): void {
        let existingItemIndex = Utils_Array.findIndex(this.items, (existingItem: IBugBashItem) => Utils_String.equals(item.id, existingItem.id, true));
        if (existingItemIndex != -1) {
            // Overwrite the item data
            this.items[existingItemIndex] = item;
        }
        else {
            this.items.push(item);
        }
    }

    private _removeItem(item: IBugBashItem): void {
        const existingItemIndex = Utils_Array.findIndex(this.items, (existingItem: IBugBashItem) => Utils_String.equals(item.id, existingItem.id, true));

        if (existingItemIndex != -1) {
            this.items.splice(existingItemIndex, 1);
        }
    }

    private _translateDates(item: IBugBashItem) {
        if (typeof item.createdDate === "string") {
            if ((item.createdDate as string).trim() === "") {
                item.createdDate = undefined;
            }
            else {
                item.createdDate = new Date(item.createdDate);
            }
        }
    }
}