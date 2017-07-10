import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");

import { BaseStore } from "VSTS_Extension/Flux/Stores/BaseStore";
import { IBugBashItem } from "../Interfaces";
import { BugBashItemActionsCreator } from "../Actions/ActionsCreator";

export interface BugBashItemStoreKey {
    bugBashId: string;
    bugBashItemId?: string;
}

export class BugBashItemStore extends BaseStore<IDictionaryStringTo<IBugBashItem[]>, IBugBashItem[], string> {
    constructor() {
        super();
        this.items = {};    
    }

    public getItem(bugBashId: string): IBugBashItem[] {
         return this.getBugBashItems(bugBashId);
    }

    public getBugBashItem(bugBashId: string, bugBashItemId: string): IBugBashItem {
         const items = this.getBugBashItems(bugBashItemId);
         if (items) {
            return Utils_Array.first(items, (item: IBugBashItem) => Utils_String.equals(item.id, bugBashItemId, true));
         }
         
         return null;
    }

    public getBugBashItems(bugBashItemId: string): IBugBashItem[] {
         return this.items[bugBashItemId.toLowerCase()] || null;
    }

    protected initializeActionListeners() {
        BugBashItemActionsCreator.InitializeItems.addListener((bugBashItems: {bugBashId: string, bugBashItems: IBugBashItem[]}) => {
            if (bugBashItems) {
                this.items[bugBashItems.bugBashId.toLowerCase()] = bugBashItems.bugBashItems;
            }

            this.emitChanged();
        });

        BugBashItemActionsCreator.RefreshItems.addListener((bugBashItems: {bugBashId: string, bugBashItems: IBugBashItem[]}) => {
            this.items[bugBashItems.bugBashId.toLowerCase()] = bugBashItems.bugBashItems;
            this.emitChanged();
        });

        BugBashItemActionsCreator.RefreshItem.addListener((bugBashItem: {bugBashId: string, bugBashItem: IBugBashItem}) => {
            this._addItem(bugBashItem.bugBashId, bugBashItem.bugBashItem);
            this.emitChanged();
        });

        BugBashItemActionsCreator.CreateItem.addListener((bugBashItem: {bugBashId: string, bugBashItem: IBugBashItem}) => {
            this._addItem(bugBashItem.bugBashId, bugBashItem.bugBashItem);
            this.emitChanged();
        });

        BugBashItemActionsCreator.UpdateItem.addListener((bugBashItem: {bugBashId: string, bugBashItem: IBugBashItem}) => {
            this._addItem(bugBashItem.bugBashId, bugBashItem.bugBashItem);
            this.emitChanged();
        });

        BugBashItemActionsCreator.AcceptItem.addListener((bugBashItem: {bugBashId: string, bugBashItem: IBugBashItem}) => {
            this._addItem(bugBashItem.bugBashId, bugBashItem.bugBashItem);
            this.emitChanged();
        });
    } 

    public getKey(): string {
        return "BugBashItemStore";
    }

    protected convertItemKeyToString(key: string): string {
        return key;
    }

    private _addItem(bugBashId: string, bugBashItem: IBugBashItem): void {
        if (!bugBashItem) {
            return;
        }

        if (this.items[bugBashId.toLowerCase()] == null) {
            this.items[bugBashId.toLowerCase()] = [];
        }

        const existingItemIndex = Utils_Array.findIndex(this.items[bugBashId.toLowerCase()], (existingItem: IBugBashItem) => Utils_String.equals(bugBashItem.id, existingItem.id, true));
        if (existingItemIndex !== -1) {
            // Overwrite the item data
            this.items[bugBashId.toLowerCase()][existingItemIndex] = bugBashItem;
        }
        else {
            this.items[bugBashId.toLowerCase()].push(bugBashItem);
        }
    }

    private _removeItem(bugBashId: string, bugBashItem: IBugBashItem): void {
        if (!bugBashItem || this.items[bugBashId.toLowerCase()] == null || this.items[bugBashId.toLowerCase()].length === 0) {
            return;
        }

        const existingItemIndex = Utils_Array.findIndex(this.items[bugBashId.toLowerCase()], (existingItem: IBugBashItem) => Utils_String.equals(bugBashItem.id, existingItem.id, true));

        if (existingItemIndex !== -1) {
            this.items[bugBashId.toLowerCase()].splice(existingItemIndex, 1);
        }
    }
}