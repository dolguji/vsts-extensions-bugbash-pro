import { StringUtils } from "MB/Utils/String";
import { ArrayUtils } from "MB/Utils/Array";
import { BaseStore } from "MB/Flux/Stores/BaseStore";

import { IBugBashItem, IBugBashItemActionData, IBugBashItemIdActionData, IBugBashItemsActionData } from "../Interfaces";
import { BugBashItemActionsHub } from "../Actions/ActionsHub";
import { BugBashItem } from "../ViewModels/BugBashItem";

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
            this._refreshBugBashItems(data);
            this.emitChanged();
        });

        BugBashItemActionsHub.RefreshBugBashItems.addListener((data: IBugBashItemsActionData) => {
            this._refreshBugBashItems(data);
            this.emitChanged();
        });

        BugBashItemActionsHub.RefreshBugBashItem.addListener((data: IBugBashItemActionData) => {
            if (data) {
                this._addBugBashItem(data.bugBashId, data.bugBashItemModel);
            }
            this.emitChanged();
        });

        BugBashItemActionsHub.CreateBugBashItem.addListener((data: IBugBashItemActionData) => {
            if (data) {
                this._newBugBashItem.reset(false);
                this._addBugBashItem(data.bugBashId, data.bugBashItemModel);
            }            
            this.emitChanged();
        });

        BugBashItemActionsHub.UpdateBugBashItem.addListener((data: IBugBashItemActionData) => {
            if (data) {
                this._addBugBashItem(data.bugBashId, data.bugBashItemModel);
            }
            this.emitChanged();
        });

        BugBashItemActionsHub.DeleteBugBashItem.addListener((data: IBugBashItemIdActionData) => {
            if (data) {
                this._removeBugBashItem(data.bugBashId, data.bugBashItemId);
            }
            this.emitChanged();
        });

        BugBashItemActionsHub.AcceptBugBashItem.addListener((data: IBugBashItemActionData) => {
            if (data) {
                this._addBugBashItem(data.bugBashId, data.bugBashItemModel);
            }
            this.emitChanged();
        });        
    } 

    public getKey(): string {
        return "BugBashItemStore";
    }

    protected convertItemKeyToString(key: string): string {
        return key;
    }

    private _refreshBugBashItems(data: IBugBashItemsActionData) {
        if (data && data.bugBashId && data.bugBashItemModels) {
            const bugBashId = data.bugBashId;
            const bugBashItemModels = data.bugBashItemModels;

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
        const existingIndex = ArrayUtils.findIndex(this.items[bugBashId.toLowerCase()], (existingBugBashItem: BugBashItem) => StringUtils.equals(bugBashItemModel.id, existingBugBashItem.id, true));
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

        const existingBugBashItemIndex = ArrayUtils.findIndex(this.items[bugBashId.toLowerCase()], (existingBugBashItem: BugBashItem) => StringUtils.equals(bugBashItemId, existingBugBashItem.id, true));

        if (existingBugBashItemIndex !== -1) {
            this.items[bugBashId.toLowerCase()].splice(existingBugBashItemIndex, 1);
        }
    }
}