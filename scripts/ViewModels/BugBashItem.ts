import { IBugBashItem } from "../Interfaces";
import { BugBashItemActions } from "../Actions/BugBashItemActions";
import { StoresHub } from "../Stores/StoresHub";

import Utils_String = require("VSS/Utils/String");

export class BugBashItem {
    public static getNewBugBashItemModel(bugBashId?: string): IBugBashItem {
        return {
            id: "",
            bugBashId: bugBashId,
            __etag: 0,
            title: "",
            description: "",
            teamId: "",
            workItemId: 0,
            createdDate: null,
            createdBy: "",
            rejected: false,
            rejectReason: "",
            rejectedBy: ""
        };
    }

    private _originalModel: IBugBashItem;
    private _updates: IBugBashItem;
    private _newComment: string;

    get newComment(): string {
        return this._newComment;
    }

    get id(): string {
        return this._originalModel.id;
    }

    get bugBashId(): string {
        return this._originalModel.bugBashId;
    }

    get workItemId(): number {
        return this._originalModel.workItemId;
    }

    get version(): number {
        return this._originalModel.__etag;
    }

    constructor(model?: IBugBashItem) {
        const bugBashItemModel = model || BugBashItem.getNewBugBashItemModel();
        this._originalModel = {...bugBashItemModel};
        this._updatedModel = {...bugBashItemModel};
        this._newComment = "";
    }

    public update(model: IBugBashItem, newComment?: string, fireChange: boolean = true) {
        this._updatedModel = {...model};

        if (newComment != null) {
            this._newComment = newComment;
        }
        
        if (fireChange) {
            BugBashItemActions.fireStoreChange();
        }        
    }

    public async save(bugBashId: string) {
        if (this.isDirty() && this.isValid()) {
            if (this.isNew()) {
                BugBashItemActions.createBugBashItem(bugBashId, this._updatedModel);
            }
            else {
                BugBashItemActions.updateBugBashItem(this.bugBashId, this._updatedModel);
            }
        }
    }

    public reset(fireChange: boolean = true) {
        if (this.isDirty()) {
            this.update(this._originalModel, "", fireChange);
        }
    }

    public delete() {
        if (!this.isNew()) {
            BugBashItemActions.deleteBugBashItem(this.bugBashId, this.id);
        }        
    }

    public async refresh() {
        if (!this.isNew()) {
            BugBashItemActions.refreshItem(this.bugBashId, this.id);  
        }
    }

    public async accept() {
        if (!this.isDirty() && !this.isNew()) {
            BugBashItemActions.acceptBugBashItem(this.bugBashId, this._originalModel); 
        }
    }

    public isNew(): boolean {
        return this.id == null || this.id.trim() === "";
    }
    
    public isAccepted(): boolean {
        return this.workItemId != null && this.workItemId > 0;
    }

    public isDirty(): boolean {
        const updatedModel = this._updatedModel;
        const originalModel = this._originalModel
        const newComment = this._newComment;

        return !Utils_String.equals(updatedModel.title, originalModel.title)
            || !Utils_String.equals(updatedModel.teamId, originalModel.teamId)
            || !Utils_String.equals(updatedModel.description, originalModel.description)
            || !Utils_String.equals(updatedModel.rejectReason, originalModel.rejectReason)
            || Boolean(updatedModel.rejected) !== Boolean(originalModel.rejected)
            || (newComment != null && newComment.trim() !== "");
    }

    public isValid(): boolean {       
        const updatedModel = this._updatedModel;

        let dataValid = updatedModel.title.trim().length > 0 
            && updatedModel.title.trim().length <= 256 
            && updatedModel.teamId.trim().length > 0
            && (!updatedModel.rejected || (updatedModel.rejectReason != null && updatedModel.rejectReason.trim().length > 0 && updatedModel.rejectReason.trim().length <= 128));

        if (dataValid) {
            dataValid = dataValid && StoresHub.teamStore.getItem(updatedModel.teamId) != null;
        }

        return dataValid;
    }
}