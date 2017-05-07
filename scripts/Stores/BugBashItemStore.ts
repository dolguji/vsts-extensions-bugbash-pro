import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");
import { WorkItemTemplate } from "TFS/WorkItemTracking/Contracts";
import * as WitClient from "TFS/WorkItemTracking/RestClient";

import { BaseStore } from "VSTS_Extension/Stores/BaseStore";

import { IBugBash } from "../Models";
import { BugBashManager } from "../BugbashManager";

export class BugBashItemStore extends BaseStore<IBugBash[], IBugBash, string> {
    constructor() {
        super();
        this.items = null;    
    }

    protected getItemByKey(id: string): IBugBash {
         return Utils_Array.first(this.items, (item: IBugBash) => Utils_String.equals(item.id, id, true));
    }

    protected async initializeItems(): Promise<void> {
        this.items = await BugBashManager.readBugBashes();
    }

    public getKey(): string {
        return "WorkItemBugBashItemStore";
    }    

    public async ensureBugBash(id: string): Promise<boolean> {
        if (!this.itemExists(id)) {
            try {
                const bugbash = await BugBashManager.readBugBash(id);
                if (bugbash) {
                    this.addItems(bugbash);
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

    public async addOrUpdateItem(bugBash: IBugBash): Promise<IBugBash> {
        bugBash.id = bugBash.id || Date.now().toString();    
        const savedBugBash = await BugBashManager.writeBugBash(bugBash);

        if (savedBugBash) {
            this.addItems(savedBugBash);
        }

        return savedBugBash;
    }

    public async deleteBugBash(bugBash: IBugBash): Promise<boolean> {         
        const deleted = await BugBashManager.deleteBugBash(bugBash);

        if (deleted) {
            this.removeItems(bugBash);
        }

        return deleted;
    }

    public async refreshItems() {
        if (this.isLoaded()) {
            // refresh only if the store has already been laoded
            this.items = null;
            this.emitChanged();
            
            await this.initializeItems();
            this.emitChanged();
        }
    }

    public addItems(items: IBugBash | IBugBash[]): void {
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

    public removeItems(items: IBugBash | IBugBash[]): void {
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
}