import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");
import { WorkItemTemplate } from "TFS/WorkItemTracking/Contracts";
import * as WitClient from "TFS/WorkItemTracking/RestClient";

import { BaseStore } from "VSTS_Extension/Stores/BaseStore";
import { ExtensionDataManager } from "VSTS_Extension/Utilities/ExtensionDataManager";

import { IBugBash } from "../Models";

export class BugBashStore extends BaseStore<IBugBash[], IBugBash, string> {
    protected getItemByKey(id: string): IBugBash {
         return Utils_Array.first(this.items, (item: IBugBash) => Utils_String.equals(item.id, id, true));
    }

    protected async initializeItems(): Promise<void> {
        let bugBashes = await ExtensionDataManager.readDocuments<IBugBash>("bugbashes", false);
        for(let bugBash of bugBashes) {
            this._translateDates(bugBash);
        }

        this.items = bugBashes;
    }

    public getKey(): string {
        return "WorkItemBugBashItemStore";
    }    

    public async ensureItem(id: string): Promise<boolean> {
        if (!this.itemExists(id)) {
            try {
                let bugBash = await ExtensionDataManager.readDocument<IBugBash>("bugbashes", id, null, false);
                if (bugBash) {
                    this._translateDates(bugBash);
                    this._addItems(bugBash);
                    return true;
                }
                else {
                    return false;
                }
            }
            catch (e) {
                console.log(e);
                return false;
            }            
        }
        else {
            this.emitChanged();
            return true;
        }
    }

    public async updateItem(bugBash: IBugBash): Promise<IBugBash> {
        try {
            const savedBugBash = await ExtensionDataManager.updateDocument<IBugBash>("bugbashes", bugBash, false);
            if (savedBugBash) {
                this._translateDates(savedBugBash);
                this._addItems(savedBugBash);
            }
            return savedBugBash;
        }
        catch (e) {
            return null;
        }
    }

    public async createItem(bugBash: IBugBash): Promise<IBugBash> {
        let model = {...bugBash};
        model.id = model.id || Date.now().toString();
        try {
            const savedBugBash = await ExtensionDataManager.createDocument<IBugBash>("bugbashes", model, false);

            if (savedBugBash) {
                this._translateDates(savedBugBash);
                this._addItems(savedBugBash);
            }

            return savedBugBash;
        }
        catch (e) {
            return null;
        }
    }

    public async deleteItem(bugBash: IBugBash): Promise<void> {
        this._removeItems(bugBash);
        await ExtensionDataManager.deleteDocument<IBugBash>("bugbashes", bugBash.id, false);
    }

    public async refreshItems() {
        if (this.isLoaded()) {
            await this.initializeItems();
            this.emitChanged();
        }
    }

    private _addItems(items: IBugBash | IBugBash[]): void {
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

    private _removeItems(items: IBugBash | IBugBash[]): void {
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

    private _addItem(item: IBugBash): void {
        const existingItemIndex = Utils_Array.findIndex(this.items, (existingItem: IBugBash) => Utils_String.equals(item.id, existingItem.id, true));
        if (existingItemIndex != -1) {
            // Overwrite the item data
            this.items[existingItemIndex] = item;
        }
        else {
            this.items.push(item);
        }
    }

    private _removeItem(item: IBugBash): void {
        const existingItemIndex = Utils_Array.findIndex(this.items, (existingItem: IBugBash) => Utils_String.equals(item.id, existingItem.id, true));

        if (existingItemIndex != -1) {
            this.items.splice(existingItemIndex, 1);
        }
    }

    private _translateDates(item: IBugBash) {
        if (typeof item.startTime === "string") {
            if ((item.startTime as string).trim() === "") {
                item.startTime = undefined;
            }
            else {
                item.startTime = new Date(item.startTime);
            }
        }
        if (typeof item.endTime === "string") {
            if ((item.endTime as string).trim() === "") {
                item.endTime = undefined;
            }
            else {
                item.endTime = new Date(item.endTime);
            }
        }
    }
}