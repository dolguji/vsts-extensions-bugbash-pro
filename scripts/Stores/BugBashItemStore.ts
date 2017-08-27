import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");

import { BaseStore } from "VSTS_Extension/Flux/Stores/BaseStore";
import { IBugBashItem, IBugBashItemActionData, IBugBashItemIdActionData, IBugBashItemsActionData } from "../Interfaces";
import { BugBashItemActionsHub } from "../Actions/ActionsHub";
import { BugBashItem } from "scripts/ViewModels/BugBashItem";

export interface BugBashItemStoreKey {
    bugBashId: string;
    bugBashItemId?: string;
}

export class BugBashItemStore extends BaseStore<IDictionaryStringTo<BugBashItem[]>, BugBashItem[], string> {
    private _itemsIdMap: IDictionaryStringTo<IDictionaryStringTo<BugBashItem>>;
    private _newBugBashItem: BugBashItem;

    constructor() {
        super();
        this.items = {};
        this._itemsIdMap = {};  
        this._newBugBashItem = new BugBashItem();
    }

    public getNewBugBashItem(): BugBashItem {
        return this._newBugBashItem;
    }

    public getItem(bugBashId: string): BugBashItem[] {
         return this.getBugBashItems(bugBashId);
    }

    public getBugBashItems(bugBashId: string): BugBashItem[] {
         return this.items[bugBashId.toLowerCase()];
    }

    public getBugBashItem(bugBashId: string, bugBashItemId: string): BugBashItem {
        if (this._itemsIdMap[bugBashId.toLowerCase()]) {
            return this._itemsIdMap[bugBashId.toLowerCase()][bugBashItemId.toLowerCase()];
        }

        return null;         
   }

    protected initializeActionListeners() {
        BugBashItemActionsHub.FireStoreChange.addListener(() => {
            this.emitChanged();
        });        

        BugBashItemActionsHub.InitializeBugBashItems.addListener((data: IBugBashItemsActionData) => {
            this._refreshBugBashItems(data.bugBashId, data.bugBashItemModels);
            this.emitChanged();
        });

        BugBashItemActionsHub.RefreshBugBashItems.addListener((data: IBugBashItemsActionData) => {
            this._refreshBugBashItems(data.bugBashId, data.bugBashItemModels);
            this.emitChanged();
        });

        BugBashItemActionsHub.RefreshBugBashItem.addListener((data: IBugBashItemActionData) => {
            this._addBugBashItem(data.bugBashId, data.bugBashItemModel);
            this.emitChanged();
        });

        BugBashItemActionsHub.CreateBugBashItem.addListener((data: IBugBashItemActionData) => {
            this._newBugBashItem.reset(false);
            this._addBugBashItem(data.bugBashId, data.bugBashItemModel);
            this.emitChanged();
        });

        BugBashItemActionsHub.UpdateBugBashItem.addListener((data: IBugBashItemActionData) => {
            this._addBugBashItem(data.bugBashId, data.bugBashItemModel);
            this.emitChanged();
        });

        BugBashItemActionsHub.DeleteBugBashItem.addListener((data: IBugBashItemIdActionData) => {
            this._removeBugBashItem(data.bugBashId, data.bugBashItemId);
            this.emitChanged();
        });

        BugBashItemActionsHub.AcceptBugBashItem.addListener((data: IBugBashItemActionData) => {
            this._addBugBashItem(data.bugBashId, data.bugBashItemModel);
            this.emitChanged();
        });        
    } 

    public getKey(): string {
        return "BugBashItemStore";
    }

    protected convertItemKeyToString(key: string): string {
        return key;
    }

    private _refreshBugBashItems(bugBashId: string, bugBashItemModels: IBugBashItem[]) {
        if (bugBashItemModels) {
            this.items[bugBashId.toLowerCase()] = [];
            this._itemsIdMap[bugBashId.toLowerCase()] = {};
            for (const bugBashItemModel of bugBashItemModels) {
                const bugBashItem = new BugBashItem(bugBashItemModel);
                this.items[bugBashId.toLowerCase()].push(bugBashItem);
                this._itemsIdMap[bugBashId.toLowerCase()][bugBashItemModel.id.toLowerCase()] = bugBashItem;
            }
        }
    }

    private _addBugBashItem(bugBashId: string, bugBashItemModel: IBugBashItem): void {
        if (!bugBashItemModel || !bugBashId) {
            return;
        }

        if (this.items[bugBashId.toLowerCase()] == null) {
            this.items[bugBashId.toLowerCase()] = [];
        }
        if (this._itemsIdMap[bugBashId.toLowerCase()] == null) {
            this._itemsIdMap[bugBashId.toLowerCase()] = {};
        }

        const bugBashItem = new BugBashItem(bugBashItemModel);
        this._itemsIdMap[bugBashId.toLowerCase()][bugBashItemModel.id.toLowerCase()] = bugBashItem;
        const existingIndex = Utils_Array.findIndex(this.items[bugBashId.toLowerCase()], (existingBugBashItem: BugBashItem) => Utils_String.equals(bugBashItemModel.id, existingBugBashItem.id, true));
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

        if (this._itemsIdMap[bugBashId.toLowerCase()] != null) {
            delete this._itemsIdMap[bugBashId.toLowerCase()][bugBashItemId.toLowerCase()];
        }

        const existingBugBashItemIndex = Utils_Array.findIndex(this.items[bugBashId.toLowerCase()], (existingBugBashItem: BugBashItem) => Utils_String.equals(bugBashItemId, existingBugBashItem.id, true));

        if (existingBugBashItemIndex !== -1) {
            this.items[bugBashId.toLowerCase()].splice(existingBugBashItemIndex, 1);
        }
    }
}