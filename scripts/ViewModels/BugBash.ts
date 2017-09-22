import { IBugBash } from "../Interfaces";
import { BugBashActions } from "../Actions/BugBashActions";
import { BugBashFieldNames, SizeLimits } from "../Constants";
import { StoresHub } from "../Stores/StoresHub";

import { StringUtils } from "MB/Utils/String";
import { DateUtils } from "MB/Utils/Date";
import { FieldType } from "TFS/WorkItemTracking/Contracts";

export class BugBash {
    public static getNewBugBashModel(): IBugBash {
        return {
            title: "New Bug Bash",
            projectId: VSS.getWebContext().project.id,
            workItemType: "",
            itemDescriptionField: "",
            autoAccept: false,
            acceptTemplateTeam: VSS.getWebContext().team.id,
            acceptTemplateId: ""
        };
    }

    private _originalModel: IBugBash;
    private _updates: IBugBash;

    get id(): string {
        return this._originalModel.id;
    }

    get projectId(): string {
        return this._originalModel.projectId;
    }

    get version(): number {
        return this._originalModel.__etag;
    }

    get isAutoAccept(): boolean {
        return this._originalModel.autoAccept;
    }

    constructor(model?: IBugBash) {
        const bugBashModel = model || BugBash.getNewBugBashModel();
        this._originalModel = {...bugBashModel};
        this._updates = {} as IBugBash;
    }

    public setFieldValue<T>(fieldName: BugBashFieldNames, fieldValue: T, fireChange: boolean = true) {
        this._updates[fieldName] = fieldValue;

        if (fireChange) {
            BugBashActions.fireStoreChange();
        }        
    }

    public getFieldValue<T>(fieldName: BugBashFieldNames, original?: boolean): T {
        if (original) {
            return this._originalModel[fieldName] as T;
        }
        else {
            const updatedModel: IBugBash = {...this._originalModel, ...this._updates};
            return updatedModel[fieldName] as T;
        }
    }

    public save() {
        if (this.isDirty() && this.isValid()) {
            const updatedModel: IBugBash = {...this._originalModel, ...this._updates};
            if (this.isNew()) {
                BugBashActions.createBugBash(updatedModel);
            }
            else {
                BugBashActions.updateBugBash(updatedModel);
            }
        }
    }

    public reset(fireChange: boolean = true) {
        this._updates = {} as IBugBash;
        if (fireChange) {
            BugBashActions.fireStoreChange();
        }
    }

    public refresh() {
        if (!this.isNew()) {
            BugBashActions.refreshBugBash(this.id);
        }
    }

    public delete() {
        if (!this.isNew()) {
            BugBashActions.deleteBugBash(this.id);
        }        
    }

    public isNew(): boolean {
        return this.id == null || this.id.trim() === "";
    }

    public isDirty(): boolean {
        const updatedModel: IBugBash = {...this._originalModel, ...this._updates};

        return !StringUtils.equals(updatedModel.title, this._originalModel.title)
            || !StringUtils.equals(updatedModel.workItemType, this._originalModel.workItemType, true)
            || !DateUtils.equals(updatedModel.startTime, this._originalModel.startTime)
            || !DateUtils.equals(updatedModel.endTime, this._originalModel.endTime)
            || !StringUtils.equals(updatedModel.itemDescriptionField, this._originalModel.itemDescriptionField, true)
            || updatedModel.autoAccept !== this._originalModel.autoAccept
            || !StringUtils.equals(updatedModel.acceptTemplateTeam, this._originalModel.acceptTemplateTeam, true)
            || !StringUtils.equals(updatedModel.acceptTemplateId, this._originalModel.acceptTemplateId, true);

    }

    public isValid(): boolean {
        const updatedModel: IBugBash = {...this._originalModel, ...this._updates};

        let dataValid = updatedModel.title.trim().length > 0
            && updatedModel.title.length <= SizeLimits.TitleFieldMaxLength
            && updatedModel.workItemType.trim().length > 0
            && updatedModel.itemDescriptionField.trim().length > 0
            && updatedModel.acceptTemplateId.trim().length > 0
            && updatedModel.acceptTemplateTeam.trim().length > 0
            && (!updatedModel.startTime || !updatedModel.endTime || DateUtils.defaultComparer(updatedModel.startTime, updatedModel.endTime) < 0);

        if (dataValid && StoresHub.teamStore.isLoaded() && StoresHub.workItemTypeStore.isLoaded() && StoresHub.workItemFieldStore.isLoaded()) {
            dataValid = dataValid 
                && StoresHub.teamStore.getItem(updatedModel.acceptTemplateTeam) != null
                && StoresHub.workItemTypeStore.getItem(updatedModel.workItemType) != null
                && StoresHub.workItemFieldStore.getItem(updatedModel.itemDescriptionField) != null
                && StoresHub.workItemFieldStore.getItem(updatedModel.itemDescriptionField).type === FieldType.Html
        }

        return dataValid;
    }
}