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
         const bugBashItems = this.getBugBashItems(bugBashId);
         if (bugBashItems) {
            return Utils_Array.first(bugBashItems, (bugBashItem: IBugBashItem) => Utils_String.equals(bugBashItem.id, bugBashItemId, true));
         }
         
         return null;
    }

    public getBugBashItems(bugBashId: string): IBugBashItem[] {
         return this.items[bugBashId.toLowerCase()] || null;
    }

    protected initializeActionListeners() {
        BugBashItemActionsCreator.InitializeBugBashItems.addListener((bugBashItems: {bugBashId: string, bugBashItems: IBugBashItem[]}) => {
            if (bugBashItems) {
                this.items[bugBashItems.bugBashId.toLowerCase()] = bugBashItems.bugBashItems;
            }

            this.emitChanged();
        });

        BugBashItemActionsCreator.RefreshBugBashItems.addListener((bugBashItems: {bugBashId: string, bugBashItems: IBugBashItem[]}) => {
            this.items[bugBashItems.bugBashId.toLowerCase()] = bugBashItems.bugBashItems;
            this.emitChanged();
        });

        BugBashItemActionsCreator.RefreshBugBashItem.addListener((bugBashItem: {bugBashId: string, bugBashItem: IBugBashItem}) => {
            this._addBugBashItem(bugBashItem.bugBashId, bugBashItem.bugBashItem);
            this.emitChanged();
        });

        BugBashItemActionsCreator.CreateBugBashItem.addListener((bugBashItem: {bugBashId: string, bugBashItem: IBugBashItem}) => {
            this._addBugBashItem(bugBashItem.bugBashId, bugBashItem.bugBashItem);
            this.emitChanged();
        });

        BugBashItemActionsCreator.UpdateBugBashItem.addListener((bugBashItem: {bugBashId: string, bugBashItem: IBugBashItem}) => {
            this._addBugBashItem(bugBashItem.bugBashId, bugBashItem.bugBashItem);
            this.emitChanged();
        });

        BugBashItemActionsCreator.DeleteBugBashItem.addListener((bugBashItem: {bugBashId: string, bugBashItemId: string}) => {
            this._removeBugBashItem(bugBashItem.bugBashId, bugBashItem.bugBashItemId);
            this.emitChanged();
        });

        BugBashItemActionsCreator.AcceptBugBashItem.addListener((bugBashItem: {bugBashId: string, bugBashItem: IBugBashItem}) => {
            this._addBugBashItem(bugBashItem.bugBashId, bugBashItem.bugBashItem);
            this.emitChanged();
        });
    } 

    public getKey(): string {
        return "BugBashItemStore";
    }

    protected convertItemKeyToString(key: string): string {
        return key;
    }

    private _addBugBashItem(bugBashId: string, bugBashItem: IBugBashItem): void {
        if (!bugBashItem) {
            return;
        }

        if (this.items[bugBashId.toLowerCase()] == null) {
            this.items[bugBashId.toLowerCase()] = [];
        }

        const existingBugBashItemIndex = Utils_Array.findIndex(this.items[bugBashId.toLowerCase()], (existingBugBashItem: IBugBashItem) => Utils_String.equals(bugBashItem.id, existingBugBashItem.id, true));
        if (existingBugBashItemIndex !== -1) {
            this.items[bugBashId.toLowerCase()][existingBugBashItemIndex] = bugBashItem;
        }
        else {
            this.items[bugBashId.toLowerCase()].push(bugBashItem);
        }
    }

    private _removeBugBashItem(bugBashId: string, bugBashItemId: string): void {
        if (!bugBashItemId || this.items[bugBashId.toLowerCase()] == null || this.items[bugBashId.toLowerCase()].length === 0) {
            return;
        }

        const existingBugBashItemIndex = Utils_Array.findIndex(this.items[bugBashId.toLowerCase()], (existingBugBashItem: IBugBashItem) => Utils_String.equals(bugBashItemId, existingBugBashItem.id, true));

        if (existingBugBashItemIndex !== -1) {
            this.items[bugBashId.toLowerCase()].splice(existingBugBashItemIndex, 1);
        }
    }
}