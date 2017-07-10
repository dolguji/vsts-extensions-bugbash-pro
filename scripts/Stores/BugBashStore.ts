import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");

import { BaseStore } from "VSTS_Extension/Flux/Stores/BaseStore";
import { IBugBash } from "../Interfaces";
import { BugBashActionsCreator } from "../Actions/ActionsCreator";

export class BugBashStore extends BaseStore<IBugBash[], IBugBash, string> {
    public getItem(id: string): IBugBash {
         return Utils_Array.first(this.items || [], (item: IBugBash) => Utils_String.equals(item.id, id, true));
    }

    protected initializeActionListeners() {
        BugBashActionsCreator.InitializeAllBugBashes.addListener((bugBashes: IBugBash[]) => {
            if (bugBashes) {
                this.items = bugBashes;
            }

            this.emitChanged();
        }); 

        BugBashActionsCreator.RefreshAllBugBashes.addListener((bugBashes: IBugBash[]) => {
            this.items = bugBashes;
            this.emitChanged();
        });

        BugBashActionsCreator.InitializeBugBash.addListener((bugBash: IBugBash) => {
            if (bugBash) {
                this._addItem(bugBash);
            }

            this.emitChanged();
        });  

        BugBashActionsCreator.CreateBugBash.addListener((bugBash: IBugBash) => {
            this._addItem(bugBash);
            this.emitChanged();
        });  

        BugBashActionsCreator.DeleteBugBash.addListener((bugBash: IBugBash) => {
            this._removeItem(bugBash);
            this.emitChanged();
        });

        BugBashActionsCreator.UpdateBugBash.addListener((bugBash: IBugBash) => {
            this._addItem(bugBash);
            this.emitChanged();
        });   
    }

    public getKey(): string {
        return "BugBashStore";
    }

    protected convertItemKeyToString(key: string): string {
        return key;
    }    
    
    private _addItem(item: IBugBash): void {
        if (!item) {
            return;
        }

        if (!this.items) {
            this.items = [];
        }

        const existingItemIndex = Utils_Array.findIndex(this.items, (existingItem: IBugBash) => Utils_String.equals(item.id, existingItem.id, true));
        if (existingItemIndex !== -1) {
            // Overwrite the item data
            this.items[existingItemIndex] = item;
        }
        else {
            this.items.push(item);
        }
    }

    private _removeItem(item: IBugBash): void {
        if (!item || this.items == null || this.items.length === 0) {
            return;
        }

        const existingItemIndex = Utils_Array.findIndex(this.items, (existingItem: IBugBash) => Utils_String.equals(item.id, existingItem.id, true));

        if (existingItemIndex !== -1) {
            this.items.splice(existingItemIndex, 1);
        }
    }
}