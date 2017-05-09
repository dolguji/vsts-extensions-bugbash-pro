import { IBugBash } from "../Models";

import Utils_String = require("VSS/Utils/String");
import Utils_Date = require("VSS/Utils/Date");

export class BugBash {
    public static getNew(): BugBash {
        return new BugBash({
            id: "",
            title: "",
            __etag: 0,
            projectId: VSS.getWebContext().project.id,
            workItemType: "",
            itemDescriptionField: "",
            autoAccept: false,
            description: "",
            acceptTemplate: {
                team: VSS.getWebContext().team.id,
                templateId: ""
            }
        });
    }

    private _model: IBugBash;
    private _originalModel: IBugBash;
    private _onChangedDelegate: () => void;

    constructor(model: IBugBash) {
        this._model = {...model};
        this._originalModel = {...model};
    }

    public getModel(): IBugBash {
        return this._model;
    }

    public isNew(): boolean {
        return !this._model.id;
    }

    public reset() {
        this._model = {...this._originalModel};
        this.fireChanged();
    }

    public renew(newModel: IBugBash) {
        this._model = {...newModel};
        this._originalModel = {...newModel};
        this.fireChanged();
    }

    public isDirty(): boolean {        
        return !Utils_String.equals(this._model.title, this._originalModel.title)
            || !Utils_String.equals(this._model.workItemType, this._originalModel.workItemType, true)
            || !Utils_String.equals(this._model.description, this._originalModel.description)
            || !Utils_Date.equals(this._model.startTime, this._originalModel.startTime)
            || !Utils_Date.equals(this._model.endTime, this._originalModel.endTime)
            || !Utils_String.equals(this._model.itemDescriptionField, this._originalModel.itemDescriptionField, true)
            || this._model.autoAccept !== this._originalModel.autoAccept
            || !Utils_String.equals(this._model.acceptTemplate.team, this._originalModel.acceptTemplate.team)
            || !Utils_String.equals(this._model.acceptTemplate.templateId, this._originalModel.acceptTemplate.templateId)
    }

    public isValid(): boolean {
        return this._model.title.trim().length > 0
            && this._model.title.length <= 128
            && this._model.workItemType.trim().length > 0
            && this._model.itemDescriptionField.trim().length > 0
            && (!this._model.startTime || !this._model.endTime || Utils_Date.defaultComparer(this._model.startTime, this._model.endTime) < 0);
    }

    public attachChanged(handler: () => void) {
        this._onChangedDelegate = handler;
    }

    public detachChanged() {
        this._onChangedDelegate = null;
    }

    public fireChanged() {
        if (this._onChangedDelegate) {
            this._onChangedDelegate();
        }
    }

    public updateTitle(newTitle: string) {
        this._model.title = newTitle;
        this.fireChanged();
    }

    public updateWorkItemType(newType: string) {
        this._model.workItemType = newType;
        this._model.acceptTemplate = {team: VSS.getWebContext().team.id, templateId: ""};  // reset template
        this.fireChanged();
    }

    public updateDescription(newDescription: string) {
        this._model.description = newDescription;
        this.fireChanged();
    }

    public updateDescriptionField(fieldRefName: string) {
        this._model.itemDescriptionField = fieldRefName;
        this.fireChanged();
    }

    public updateStartTime(newStartTime: Date) {
        this._model.startTime = newStartTime;
        this.fireChanged();
    }

    public updateEndTime(newEndTime: Date) {        
        this._model.endTime = newEndTime;
        this.fireChanged();
    }

    public updateAutoAccept(autoAccept: boolean) {        
        this._model.autoAccept = autoAccept;
        this.fireChanged();
    }

    public updateAcceptTemplate(templateId: string) {
        this._model.acceptTemplate = {team: VSS.getWebContext().team.id, templateId: templateId};
        this.fireChanged();
    }
}