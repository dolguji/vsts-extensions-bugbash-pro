import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");

import { BaseStore } from "VSTS_Extension/Flux/Stores/BaseStore";
import { IBugBashItem } from "../Interfaces";
import { BugBashItemActionsHub } from "../Actions/ActionsHub";

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

    public getBugBashItems(bugBashId: string): IBugBashItem[] {
         return this.items[(bugBashId || "").toLowerCase()] || null;
    }

    protected initializeActionListeners() {
        BugBashItemActionsHub.InitializeBugBashItems.addListener((data: {bugBashId: string, bugBashItems: IBugBashItem[]}) => {
            if (data) {
                this.items[data.bugBashId.toLowerCase()] = data.bugBashItems;
            }

            this.emitChanged();
        });

        BugBashItemActionsHub.RefreshBugBashItems.addListener((data: {bugBashId: string, bugBashItems: IBugBashItem[]}) => {
            this.items[data.bugBashId.toLowerCase()] = data.bugBashItems;

            this.emitChanged();
        });

        BugBashItemActionsHub.RefreshBugBashItem.addListener((data: {bugBashId: string, bugBashItem: IBugBashItem}) => {
            this._addBugBashItem(data.bugBashId, data.bugBashItem);
            this.emitChanged();
        });

        BugBashItemActionsHub.CreateBugBashItem.addListener((data: {bugBashId: string, bugBashItem: IBugBashItem}) => {
            this._addBugBashItem(data.bugBashId, data.bugBashItem);
            this.emitChanged();
        });

        BugBashItemActionsHub.UpdateBugBashItem.addListener((data: {bugBashId: string, bugBashItem: IBugBashItem}) => {
            this._addBugBashItem(data.bugBashId, data.bugBashItem);
            this.emitChanged();
        });

        BugBashItemActionsHub.DeleteBugBashItem.addListener((data: {bugBashId: string, bugBashItemId: string}) => {
            this._removeBugBashItem(data.bugBashId, data.bugBashItemId);
            this.emitChanged();
        });

        BugBashItemActionsHub.AcceptBugBashItem.addListener((data: {bugBashId: string, bugBashItem: IBugBashItem}) => {
            this._addBugBashItem(data.bugBashId, data.bugBashItem);
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

        const existingIndex = Utils_Array.findIndex(this.items[bugBashId.toLowerCase()], (existingBugBashItem: IBugBashItem) => Utils_String.equals(bugBashItem.id, existingBugBashItem.id, true));
        if (existingIndex !== -1) {
            this.items[bugBashId.toLowerCase()][existingIndex] = bugBashItem;
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