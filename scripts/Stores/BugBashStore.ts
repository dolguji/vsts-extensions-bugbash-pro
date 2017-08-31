import { StringUtils } from "MB/Utils/String";
import { ArrayUtils } from "MB/Utils/Array";
import { BaseStore } from "MB/Flux/Stores/BaseStore";

import { BugBash } from "../ViewModels/BugBash";
import { BugBashActionsHub } from "../Actions/ActionsHub";
import { IBugBash } from "../Interfaces";

export class BugBashStore extends BaseStore<BugBash[], BugBash, string> {
    private _allLoaded: boolean = false;
    private _itemIdMap: IDictionaryStringTo<BugBash>;
    private _newBugBash: BugBash;

    constructor() {
        super();
        this._allLoaded = false;
        this._itemIdMap = {};
        this._newBugBash = new BugBash();
    }
    
    public isLoaded(key?: string): boolean {
        if (key) {
            return super.isLoaded();
        }
        
        return this._allLoaded && super.isLoaded();
    }

    public getNewBugBash(): BugBash {
        return this._newBugBash;
    }

    public getItem(id: string): BugBash {
        return this._itemIdMap[id.toLowerCase()];
    }

    protected initializeActionListeners() {
        BugBashActionsHub.FireStoreChange.addListener(() => {
            this.emitChanged();
        });

        BugBashActionsHub.InitializeAllBugBashes.addListener((bugBashModels: IBugBash[]) => {
            this._refreshBugBashes(bugBashModels);
            this._allLoaded = true;
            this.emitChanged();
        }); 

        BugBashActionsHub.RefreshAllBugBashes.addListener((bugBashModels: IBugBash[]) => {
            this._refreshBugBashes(bugBashModels);
            this.emitChanged();
        });

        BugBashActionsHub.InitializeBugBash.addListener((bugBashModel: IBugBash) => {
            this._addOrUpdateBugBash(bugBashModel);
            this.emitChanged();
        }); 
        
        BugBashActionsHub.RefreshBugBash.addListener((bugBashModel: IBugBash) => {
            this._addOrUpdateBugBash(bugBashModel);
            this.emitChanged();
        }); 

        BugBashActionsHub.CreateBugBash.addListener((bugBashModel: IBugBash) => {
            this._newBugBash.reset(false);
            this._addOrUpdateBugBash(bugBashModel);
            this.emitChanged();
        });  

        BugBashActionsHub.UpdateBugBash.addListener((bugBashModel: IBugBash) => {
            this._addOrUpdateBugBash(bugBashModel);
            this.emitChanged();
        });  
        
        BugBashActionsHub.DeleteBugBash.addListener((bugBashId: string) => {
            this._removeBugBash(bugBashId);
            this.emitChanged();
        });        
    }

    public getKey(): string {
        return "BugBashStore";
    }

    protected convertItemKeyToString(key: string): string {
        return key;
    }    
    
    private _refreshBugBashes(bugBashModels: IBugBash[]) {
        if (bugBashModels) {
            this.items = [];
            this._itemIdMap = {};

            for (const bugBashModel of bugBashModels) {
                const bugBash = new BugBash(bugBashModel);
                this.items.push(bugBash);
                this._itemIdMap[bugBashModel.id.toLowerCase()] = bugBash;
            }
        }
    }

    private _addOrUpdateBugBash(bugBashModel: IBugBash) {
        if (!bugBashModel) {
            return;
        }

        if (!this.items) {
            this.items = [];
        }

        const bugBash = new BugBash(bugBashModel);
        this._itemIdMap[bugBashModel.id.toLowerCase()] = bugBash;

        const existingBugBashIndex = ArrayUtils.findIndex(this.items, (existingBugBash: BugBash) => StringUtils.equals(bugBashModel.id, existingBugBash.id, true));
        if (existingBugBashIndex !== -1) {
            this.items[existingBugBashIndex] = bugBash;
        }
        else {
            this.items.push(bugBash);
        }
    }

    private _removeBugBash(bugBashId: string) {        
        if (!bugBashId || this.items == null || this.items.length === 0) {
            return;
        }

        delete this._itemIdMap[bugBashId.toLowerCase()];

        const existingBugBashIndex = ArrayUtils.findIndex(this.items, (existingBugBash: BugBash) => StringUtils.equals(bugBashId, existingBugBash.id, true));

        if (existingBugBashIndex !== -1) {
            this.items.splice(existingBugBashIndex, 1);
        }
    }
}