import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");

import { BaseStore } from "VSTS_Extension/Flux/Stores/BaseStore";
import { IBugBashItem, IBugBashItemViewModel } from "../Interfaces";
import { BugBashItemActionsCreator } from "../Actions/ActionsCreator";

export interface BugBashItemStoreKey {
    bugBashId: string;
    bugBashItemId?: string;
}

export class BugBashItemStore extends BaseStore<IDictionaryStringTo<IBugBashItemViewModel[]>, IBugBashItemViewModel[], string> {
    constructor() {
        super();
        this.items = {};    
    }

    public getItem(bugBashId: string): IBugBashItemViewModel[] {
         return this.getBugBashItemViewModels(bugBashId);
    }

    public getBugBashItemViewModel(bugBashId: string, bugBashItemId: string): IBugBashItemViewModel {
         const bugBashItemViewModels = this.getBugBashItemViewModels(bugBashId);
         if (bugBashItemViewModels) {
            return Utils_Array.first(bugBashItemViewModels, (bugBashItemViewModel: IBugBashItemViewModel) => Utils_String.equals(bugBashItemViewModel.updatedBugBashItem.id, bugBashItemId, true));
         }
         
         return null;
    }

    public getBugBashItemViewModels(bugBashId: string): IBugBashItemViewModel[] {
         return this.items[bugBashId.toLowerCase()] || null;
    }

    protected initializeActionListeners() {
        BugBashItemActionsCreator.InitializeBugBashItems.addListener((data: {bugBashId: string, bugBashItems: IBugBashItem[]}) => {
            if (data) {
                this.items[data.bugBashId.toLowerCase()] = data.bugBashItems.map(b => {
                    return {
                        updatedBugBashItem: {...b},
                        originalBugBashItem: {...b},
                        newComment: ""
                    };
                });
            }

            this.emitChanged();
        });

        BugBashItemActionsCreator.RefreshBugBashItems.addListener((data: {bugBashId: string, bugBashItems: IBugBashItem[]}) => {
            this.items[data.bugBashId.toLowerCase()] = data.bugBashItems.map(b => {
                return {
                    updatedBugBashItem: {...b},
                    originalBugBashItem: {...b},
                    newComment: ""
                };
            });

            this.emitChanged();
        });

        BugBashItemActionsCreator.RefreshBugBashItem.addListener((data: {bugBashId: string, bugBashItem: IBugBashItem}) => {
            this._addBugBashItem(data.bugBashId, data.bugBashItem);
            this.emitChanged();
        });

        BugBashItemActionsCreator.CreateBugBashItem.addListener((data: {bugBashId: string, bugBashItem: IBugBashItem}) => {
            this._addBugBashItem(data.bugBashId, data.bugBashItem);
            this.emitChanged();
        });

        BugBashItemActionsCreator.UpdateBugBashItem.addListener((data: {bugBashId: string, bugBashItem: IBugBashItem}) => {
            this._addBugBashItem(data.bugBashId, data.bugBashItem);
            this.emitChanged();
        });

        BugBashItemActionsCreator.DeleteBugBashItem.addListener((data: {bugBashId: string, bugBashItemId: string}) => {
            this._removeBugBashItem(data.bugBashId, data.bugBashItemId);
            this.emitChanged();
        });

        BugBashItemActionsCreator.AcceptBugBashItem.addListener((data: {bugBashId: string, bugBashItem: IBugBashItem}) => {
            this._addBugBashItem(data.bugBashId, data.bugBashItem);
            this.emitChanged();
        });

        BugBashItemActionsCreator.DirtyUpdateBugBashItem.addListener((data: {bugBashId: string, bugBashItem: IBugBashItem, newComment: string}) => {
            const bugBashId = data.bugBashId;
            const bugBashItem = data.bugBashItem;
            if (bugBashItem) {
                const existingItem = this.getBugBashItemViewModel(bugBashId, bugBashItem.id);
                if (existingItem) {                    
                    existingItem.updatedBugBashItem = {...data.bugBashItem};
                    existingItem.newComment = data.newComment;
                }
            }

            this.emitChanged();
        });

        BugBashItemActionsCreator.UndoUpdateBugBashItem.addListener((data: {bugBashId: string, bugBashItemId?: string}) => {
            const bugBashId = data.bugBashId;
            const bugBashItemId = data.bugBashItemId;
            if (bugBashItemId) {
                const existingItem = this.getBugBashItemViewModel(bugBashId, bugBashItemId);
                if (existingItem) {                    
                    existingItem.updatedBugBashItem = {...existingItem.originalBugBashItem};
                    existingItem.newComment = "";
                }
            }
            else {
                // undo all bug bash items in given bug bash
                const allItems = this.items[data.bugBashId.toLowerCase()] || [];
                for (const item of allItems) {
                    item.updatedBugBashItem = {...item.originalBugBashItem};
                    item.newComment = "";
                }
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

    private _addBugBashItem(bugBashId: string, bugBashItem: IBugBashItem): void {
        if (!bugBashItem) {
            return;
        }

        if (this.items[bugBashId.toLowerCase()] == null) {
            this.items[bugBashId.toLowerCase()] = [];
        }

        const viewModel = {
            updatedBugBashItem: {...bugBashItem},
            originalBugBashItem: {...bugBashItem},
            newComment: ""
        };
        const existingIndex = Utils_Array.findIndex(this.items[bugBashId.toLowerCase()], (existingBugBashItem: IBugBashItemViewModel) => Utils_String.equals(bugBashItem.id, existingBugBashItem.originalBugBashItem.id, true));
        if (existingIndex !== -1) {
            this.items[bugBashId.toLowerCase()][existingIndex] = viewModel;
        }
        else {
            this.items[bugBashId.toLowerCase()].push(viewModel);
        }
    }

    private _removeBugBashItem(bugBashId: string, bugBashItemId: string): void {
        if (!bugBashItemId || this.items[bugBashId.toLowerCase()] == null || this.items[bugBashId.toLowerCase()].length === 0) {
            return;
        }

        const existingBugBashItemIndex = Utils_Array.findIndex(this.items[bugBashId.toLowerCase()], (existingBugBashItem: IBugBashItemViewModel) => Utils_String.equals(bugBashItemId, existingBugBashItem.originalBugBashItem.id, true));

        if (existingBugBashItemIndex !== -1) {
            this.items[bugBashId.toLowerCase()].splice(existingBugBashItemIndex, 1);
        }
    }
}