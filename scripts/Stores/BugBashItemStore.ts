import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");
import { WorkItemTemplate } from "TFS/WorkItemTracking/Contracts";
import * as WitClient from "TFS/WorkItemTracking/RestClient";

import { ExtensionDataManager } from "VSTS_Extension/Utilities/ExtensionDataManager";
import { BaseStore } from "VSTS_Extension/Stores/BaseStore";

import { IBugBashItemDocument } from "../Models";
import { getBugBashCollectionKey } from "../Helpers";

export class BugBashItemStore extends BaseStore<IBugBashItemDocument[], IBugBashItemDocument, string> {
    private _dataLoadedMap: IDictionaryStringTo<boolean>;

    constructor() {
        super();
        this.items = [];
        this._dataLoadedMap = {};
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

    public isDataLoaded(bugBashId: string): boolean {
        return Boolean(this._dataLoadedMap[bugBashId.toLowerCase()]);
    }

    public async refreshItems(bugBashId: string) {
        delete this._dataLoadedMap[bugBashId.toLowerCase()];
        this._removeItems(this.items.filter(i => Utils_String.equals(i.bugBashId, bugBashId, true)));

        const models = await ExtensionDataManager.readDocuments<IBugBashItemDocument>(getBugBashCollectionKey(bugBashId), false);
        for(let model of models) {
            this._translateDates(model);
        }
        
        this._dataLoadedMap[bugBashId.toLowerCase()] = true;
        this._addItems(models);
    }

    public async createItem(item: IBugBashItemDocument): Promise<IBugBashItemDocument> {
        let model = {...item};
        model.id = model.id || `${model.bugBashId}_${Date.now().toString()}`;
        model.createdBy = model.createdBy || `${VSS.getWebContext().user.name} <${VSS.getWebContext().user.uniqueName}>`;
        model.createdDate = model.createdDate || new Date(Date.now());

        try {
            const savedItem = await ExtensionDataManager.createDocument(getBugBashCollectionKey(model.bugBashId), model, false);

            if (savedItem) {
                this._translateDates(savedItem);
                this._addItems(savedItem);
            }

            return savedItem;
        }
        catch (e) {
            return null;
        }

    }

    public async updateItem(item: IBugBashItemDocument): Promise<IBugBashItemDocument> {        
        try {
            const savedItem = await ExtensionDataManager.updateDocument(getBugBashCollectionKey(item.bugBashId), item, false);

            if (savedItem) {
                this._translateDates(savedItem);
                this._addItems(savedItem);
            }

            return savedItem;
        }
        catch (e) {
            return null;
        }
    }

    public async deleteItems(items: IBugBashItemDocument[]): Promise<void> {
        this._removeItems(items);
        await Promise.all(items.map(async (item) => {
            return await ExtensionDataManager.deleteDocument(getBugBashCollectionKey(item.bugBashId), item.id, false);
        }));        
    }

    public clearAllItems(bugBashId: string) {
        this.deleteItems(this.items.filter(i => Utils_String.equals(i.bugBashId, bugBashId, true)));
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

    private _translateDates(item: IBugBashItemDocument) {
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