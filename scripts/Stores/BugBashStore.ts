import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");

import { BaseStore } from "VSTS_Extension/Flux/Stores/BaseStore";
import { IBugBash } from "../Interfaces";
import { BugBashActionsHub } from "../Actions/ActionsHub";

export class BugBashStore extends BaseStore<IBugBash[], IBugBash, string> {
    private _allLoaded: boolean = false;

    public isLoaded(key?: string): boolean {
        if (key) {
            return super.isLoaded();
        }
        
        return this._allLoaded && super.isLoaded();
    }

    public getItem(id: string): IBugBash {
         return Utils_Array.first(this.items || [], (bugBash: IBugBash) => Utils_String.equals(bugBash.id, id, true));
    }

    protected initializeActionListeners() {
        BugBashActionsHub.InitializeAllBugBashes.addListener((bugBashes: IBugBash[]) => {
            if (bugBashes) {
                this.items = bugBashes;
            }
            this._allLoaded = true;
            this.emitChanged();
        }); 

        BugBashActionsHub.RefreshAllBugBashes.addListener((bugBashes: IBugBash[]) => {
            this.items = bugBashes;
            this.emitChanged();
        });

        BugBashActionsHub.InitializeBugBash.addListener((bugBash: IBugBash) => {
            if (bugBash) {
                this._addBugBash(bugBash);
            }

            this.emitChanged();
        }); 
        
        BugBashActionsHub.RefreshBugBash.addListener((bugBash: IBugBash) => {
            this._addBugBash(bugBash);
            this.emitChanged();
        }); 

        BugBashActionsHub.CreateBugBash.addListener((bugBash: IBugBash) => {
            this._addBugBash(bugBash);
            this.emitChanged();
        });  

        BugBashActionsHub.DeleteBugBash.addListener((bugBashId: string) => {
            this._removeBugBash(bugBashId);
            this.emitChanged();
        });

        BugBashActionsHub.UpdateBugBash.addListener((bugBash: IBugBash) => {
            this._addBugBash(bugBash);
            this.emitChanged();
        });  
    }

    public getKey(): string {
        return "BugBashStore";
    }

    protected convertItemKeyToString(key: string): string {
        return key;
    }    
    
    private _addBugBash(bugBash: IBugBash): void {
        if (!bugBash) {
            return;
        }

        if (!this.items) {
            this.items = [];
        }

        const existingBugBashIndex = Utils_Array.findIndex(this.items, (existingBugBash: IBugBash) => Utils_String.equals(bugBash.id, existingBugBash.id, true));
        if (existingBugBashIndex !== -1) {
            this.items[existingBugBashIndex] = bugBash;
        }
        else {
            this.items.push(bugBash);
        }
    }

    private _removeBugBash(bugBashId: string): void {
        if (!bugBashId || this.items == null || this.items.length === 0) {
            return;
        }

        const existingBugBashIndex = Utils_Array.findIndex(this.items, (existingBugBash: IBugBash) => Utils_String.equals(bugBashId, existingBugBash.id, true));

        if (existingBugBashIndex !== -1) {
            this.items.splice(existingBugBashIndex, 1);
        }
    }
}