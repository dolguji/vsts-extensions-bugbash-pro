import Utils_String = require("VSS/Utils/String");
import Utils_Date = require("VSS/Utils/Date");
import { EventHandlerList, NamedEventCollection } from "VSS/Events/Handlers";

import { IBugBashItem } from "./Interfaces";
import { StoresHub } from "./Stores/StoresHub";

export class BugBashItemProvider {
    public originalBugBashItems: IDictionaryStringTo<IBugBashItem>;
    public updatedBugBashItems: IDictionaryStringTo<IBugBashItem>;
    public bugBashItemNewComments: IDictionaryStringTo<string>;

    private _changedHandlers = new EventHandlerList();
    private _namedEventCollection = new NamedEventCollection<any, any>();

    constructor() {
        this.originalBugBashItems = {};
        this.updatedBugBashItems = {};
        this.bugBashItemNewComments = {};
    }

    public initialize(bugBashItems: IBugBashItem[]) {
        for (const bugBashItem of bugBashItems) {
            this.originalBugBashItems[bugBashItem.id] = {...bugBashItem};
            this.updatedBugBashItems[bugBashItem.id] = {...bugBashItem};
            this.bugBashItemNewComments = {};
        }
        this._emitChanged();
    }

    public dispose() {
        this.originalBugBashItems = {};
        this.updatedBugBashItems = {};
        this.bugBashItemNewComments = {};
        this._emitChanged();
    }

    public addChangedListener(handler: IEventHandler): void {
        this._changedHandlers.subscribe(handler as any);
    }

    public removeChangedListener(handler: IEventHandler): void {
        this._changedHandlers.unsubscribe(handler as any);
    }

    private _emitChanged(): void {
        this._changedHandlers.invokeHandlers(this);
    }

    private _itemExists(id: string) {
        return this.updatedBugBashItems[id] != null && this.originalBugBashItems[id] != null;
    }

    public update(bugBashItem: IBugBashItem, newComment?: string) {
        if (bugBashItem && this._itemExists(bugBashItem.id)) {
            this.updatedBugBashItems[bugBashItem.id] = {...bugBashItem};
            this.bugBashItemNewComments[bugBashItem.id] = newComment || "";
        }

        this._emitChanged();
    }    

    public renew(bugBashItem: IBugBashItem) {
        if (bugBashItem && this._itemExists(bugBashItem.id)) {
            this.updatedBugBashItems[bugBashItem.id]= {...bugBashItem};
            this.originalBugBashItems[bugBashItem.id] = {...bugBashItem};
            this.bugBashItemNewComments[bugBashItem.id] = "";
        }
        
        this._emitChanged();
    } 
    
    public undo(id: string) {
        if (this._itemExists(id)) {
            this.updatedBugBashItems[id]= {...this.originalBugBashItems[id]};
            this.bugBashItemNewComments[id] = "";
        }

        this._emitChanged();
    } 

    public isDirty(id: string): boolean {
        if (this._itemExists(id)) {
            const updatedBugBashItem = this.updatedBugBashItems[id];
            const originalBugBashItem = this.originalBugBashItems[id];
            const newComment = this.bugBashItemNewComments[id];

            return !Utils_String.equals(updatedBugBashItem.title, originalBugBashItem.title)
            || !Utils_String.equals(updatedBugBashItem.teamId, originalBugBashItem.teamId)
            || !Utils_String.equals(updatedBugBashItem.description, originalBugBashItem.description)
            || !Utils_String.equals(updatedBugBashItem.rejectReason, originalBugBashItem.rejectReason)
            || Boolean(updatedBugBashItem.rejected) !== Boolean(originalBugBashItem.rejected)
            || (newComment != null && newComment.trim() !== "");
        }

        return false;
    }

    public isAnyItemDirty(): boolean {
        const keys = Object.keys(this.updatedBugBashItems);
        for (const key of keys) {
            if (this.isDirty(this.updatedBugBashItems[key].id)) {
                return true;
            }
        }

        return false;
    }

    public isValid(id: string): boolean {
        if (this._itemExists(id)) {
            const updatedBugBashItem = this.updatedBugBashItems[id];

            let dataValid = updatedBugBashItem.title.trim().length > 0 
                && updatedBugBashItem.title.trim().length <= 256 
                && updatedBugBashItem.teamId.trim().length > 0
                && (!updatedBugBashItem.rejected || (updatedBugBashItem.rejectReason != null && updatedBugBashItem.rejectReason.trim().length > 0 && updatedBugBashItem.rejectReason.trim().length <= 128));

            if (dataValid) {
                dataValid = dataValid && StoresHub.teamStore.getItem(updatedBugBashItem.teamId) != null;
            }

            return dataValid;
        }
        
        return false;
    }
}