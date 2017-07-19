import { IBugBash } from "./Interfaces";
import Utils_String = require("VSS/Utils/String");
import Utils_Date = require("VSS/Utils/Date");
import { EventHandlerList, NamedEventCollection } from "VSS/Events/Handlers";

export class BugBashProvider {
    public originalBugBash: IBugBash;
    public updatedBugBash: IBugBash;
    private _changedHandlers = new EventHandlerList();
    private _namedEventCollection = new NamedEventCollection<any, any>();

    constructor() {}

    public initialize(bugBash: IBugBash) {
        this.originalBugBash = {...bugBash};
        this.updatedBugBash = {...bugBash};
        this._emitChanged();
    }

    public dispose() {
        this.originalBugBash = null;
        this.updatedBugBash = null;
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

    public isInitialized(): boolean {
        return this.originalBugBash != null && this.updatedBugBash != null;

    }
    
    public update(bugBash: IBugBash) {
        if (this.isInitialized()) {
            this.updatedBugBash = {...bugBash};
        }
        
        this._emitChanged();
    }    

    public renew(bugBash: IBugBash) {
        this.updatedBugBash = {...bugBash};
        this.originalBugBash = {...bugBash};
        this._emitChanged();
    } 
    
    public undo() {
        if (this.isInitialized()) {
            this.updatedBugBash = {...this.originalBugBash};
        }

        this._emitChanged();
    } 

    public isNew(): boolean {
        if (this.originalBugBash) {
            return this.originalBugBash.id == null || this.originalBugBash.id.trim() === "";
        }
        return false;
    }

    public isDirty(): boolean {
        if (this.isInitialized()) {
            const updatedBugBash = this.updatedBugBash;
            const originalBugBash = this.originalBugBash;

            return !Utils_String.equals(updatedBugBash.title, originalBugBash.title)
                || !Utils_String.equals(updatedBugBash.workItemType, originalBugBash.workItemType, true)
                || !Utils_String.equals(updatedBugBash.description, originalBugBash.description)
                || !Utils_Date.equals(updatedBugBash.startTime, originalBugBash.startTime)
                || !Utils_Date.equals(updatedBugBash.endTime, originalBugBash.endTime)
                || !Utils_String.equals(updatedBugBash.itemDescriptionField, originalBugBash.itemDescriptionField, true)
                || updatedBugBash.autoAccept !== originalBugBash.autoAccept
                || !Utils_String.equals(updatedBugBash.acceptTemplate.team, originalBugBash.acceptTemplate.team)
                || !Utils_String.equals(updatedBugBash.acceptTemplate.templateId, originalBugBash.acceptTemplate.templateId);
        }

        return false;
    }

    public isValid(): boolean {
        if (this.isInitialized()) {
            const bugBash = this.updatedBugBash;

            return bugBash.title.trim().length > 0
                && bugBash.title.length <= 256
                && bugBash.workItemType.trim().length > 0
                && bugBash.itemDescriptionField.trim().length > 0
                && (!bugBash.startTime || !bugBash.endTime || Utils_Date.defaultComparer(bugBash.startTime, bugBash.endTime) < 0);
        }
        
        return false;
    }
}