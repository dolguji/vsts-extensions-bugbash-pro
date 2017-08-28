import Utils_String = require("VSS/Utils/String");

import { IBugBashItem } from "../Interfaces";
import { BugBashItemActions } from "../Actions/BugBashItemActions";
import { StoresHub } from "../Stores/StoresHub";
import { BugBashItemFieldNames } from "../Constants";

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

    get isAccepted(): boolean {
        return this.workItemId != null && this.workItemId > 0;
    }

    constructor(model?: IBugBashItem) {
        const bugBashItemModel = model || BugBashItem.getNewBugBashItemModel();
        this._originalModel = {...bugBashItemModel};
        this._updates = {} as IBugBashItem;
        this._newComment = "";
    }

    public setFieldValue<T>(fieldName: BugBashItemFieldNames, fieldValue: T, fireChange: boolean = true) {
        this._updates[fieldName] = fieldValue;

        if (fireChange) {
            BugBashItemActions.fireStoreChange();
        }        
    }

    public getFieldValue<T>(fieldName: BugBashItemFieldNames, original?: boolean): T {
        if (original) {
            return this._originalModel[fieldName] as T;
        }
        else {
            const updatedModel: IBugBashItem = {...this._originalModel, ...this._updates};
            return updatedModel[fieldName] as T;
        }
    }

    public setComment(newComment: string) {
        this._newComment = newComment;
        BugBashItemActions.fireStoreChange();
    }

    public save(bugBashId: string) {
        if (this.isDirty() && this.isValid()) {
            const updatedModel: IBugBashItem = {...this._originalModel, ...this._updates};

            if (this.isNew()) {
                BugBashItemActions.createBugBashItem(bugBashId, updatedModel, this.newComment);
            }
            else {
                BugBashItemActions.updateBugBashItem(this.bugBashId, updatedModel, this.newComment);
            }
        }
    }

    public reset(fireChange: boolean = true) {
        this._updates = {} as IBugBashItem;
        this._newComment = "";
        if (fireChange) {
            BugBashItemActions.fireStoreChange();
        }
    }

    public delete() {
        if (!this.isNew()) {
            BugBashItemActions.deleteBugBashItem(this.bugBashId, this.id);
        }        
    }

    public refresh() {
        if (!this.isNew()) {
            BugBashItemActions.refreshItem(this.bugBashId, this.id);  
        }
    }

    public accept() {
        if (!this.isDirty() && !this.isNew()) {
            BugBashItemActions.acceptBugBashItem(this.bugBashId, this._originalModel); 
        }
    }

    public isNew(): boolean {
        return this.id == null || this.id.trim() === "";
    }    

    public isDirty(): boolean {
        const updatedModel: IBugBashItem = {...this._originalModel, ...this._updates};

        return !Utils_String.equals(updatedModel.title, this._originalModel.title)
            || !Utils_String.equals(updatedModel.teamId, this._originalModel.teamId)
            || !Utils_String.equals(updatedModel.description, this._originalModel.description)
            || !Utils_String.equals(updatedModel.rejectReason, this._originalModel.rejectReason)
            || Boolean(updatedModel.rejected) !== Boolean(this._originalModel.rejected)
            || (this._newComment != null && this._newComment.trim() !== "");
    }

    public isValid(): boolean {       
        const updatedModel: IBugBashItem = {...this._originalModel, ...this._updates};

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