import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");

import { BaseStore } from "VSTS_Extension/Flux/Stores/BaseStore";
import { IBugBash, IBugBashViewModel } from "../Interfaces";
import { BugBashActionsCreator } from "../Actions/ActionsCreator";

export class BugBashStore extends BaseStore<IBugBashViewModel[], IBugBashViewModel, string> {
    private _allLoaded: boolean = false;

    public isLoaded(key?: string): boolean {
        if (key) {
            return super.isLoaded();
        }
        
        return this._allLoaded && super.isLoaded();
    }

    public getItem(id: string): IBugBashViewModel {
         return Utils_Array.first(this.items || [], (bugBash: IBugBashViewModel) => Utils_String.equals(bugBash.originalBugBash.id, id, true));
    }

    protected initializeActionListeners() {
        BugBashActionsCreator.InitializeAllBugBashes.addListener((bugBashes: IBugBash[]) => {
            if (bugBashes) {
                this.items = bugBashes.map(b => {
                    return {
                        originalBugBash: {...b},
                        updatedBugBash: {...b}
                    }  
                });
            }
            this._allLoaded = true;
            this.emitChanged();
        }); 

        BugBashActionsCreator.RefreshAllBugBashes.addListener((bugBashes: IBugBash[]) => {
            this.items = bugBashes.map(b => {
                return {
                    originalBugBash: {...b},
                    updatedBugBash: {...b}
                }  
            });
            this.emitChanged();
        });

        BugBashActionsCreator.InitializeBugBash.addListener((bugBash: IBugBash) => {
            if (bugBash) {
                this._addBugBash(bugBash);
            }

            this.emitChanged();
        }); 
        
        BugBashActionsCreator.RefreshBugBash.addListener((bugBash: IBugBash) => {
            this._addBugBash(bugBash);
            this.emitChanged();
        }); 

        BugBashActionsCreator.CreateBugBash.addListener((bugBash: IBugBash) => {
            this._addBugBash(bugBash);
            this.emitChanged();
        });  

        BugBashActionsCreator.DeleteBugBash.addListener((bugBashId: string) => {
            this._removeBugBash(bugBashId);
            this.emitChanged();
        });

        BugBashActionsCreator.UpdateBugBash.addListener((bugBash: IBugBash) => {
            this._addBugBash(bugBash);
            this.emitChanged();
        });  

        BugBashActionsCreator.DirtyUpdateBugBash.addListener((bugBash: IBugBash) => {
            if (bugBash) {
                const existingBugBash = this.getItem(bugBash.id);                
                if (existingBugBash) {
                    existingBugBash.updatedBugBash = {...bugBash};
                }
            }

            this.emitChanged();
        });

        BugBashActionsCreator.UndoUpdateBugBash.addListener((bugBashId?: string) => {
            if (bugBashId) {
                const existingItem = this.getItem(bugBashId);
                if (existingItem) {                    
                    existingItem.updatedBugBash = {...existingItem.originalBugBash};
                }
            }
            else {
                // undo all bug bash items in given bug bash
                for (const item of (this.items || [])) {
                    item.updatedBugBash = {...item.originalBugBash};
                }
            }

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

        const viewModel = {
            updatedBugBash: {...bugBash},
            originalBugBash: {...bugBash}
        };

        const existingBugBashIndex = Utils_Array.findIndex(this.items, (existingBugBash: IBugBashViewModel) => Utils_String.equals(bugBash.id, existingBugBash.originalBugBash.id, true));
        if (existingBugBashIndex !== -1) {
            this.items[existingBugBashIndex] = viewModel;
        }
        else {
            this.items.push(viewModel);
        }
    }

    private _removeBugBash(bugBashId: string): void {
        if (!bugBashId || this.items == null || this.items.length === 0) {
            return;
        }

        const existingBugBashIndex = Utils_Array.findIndex(this.items, (existingBugBash: IBugBashViewModel) => Utils_String.equals(bugBashId, existingBugBash.originalBugBash.id, true));

        if (existingBugBashIndex !== -1) {
            this.items.splice(existingBugBashIndex, 1);
        }
    }
}