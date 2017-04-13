import {  IBugBash } from "./Models";

import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");
import Utils_Date = require("VSS/Utils/Date");
import Context = require("VSS/Context");

export class BugBash {
    public static getNew(): BugBash {
        return new BugBash({
            id: "",
            title: "",
            __etag: 0,
            templateId: "",
            manualFields: [],
            projectId: VSS.getWebContext().project.id,
            teamId: VSS.getWebContext().team.id,
            workItemType: "",
            configTemplates: {}
        });
    }

    private _model: IBugBash;
    private _originalModel: IBugBash;
    private _onChangedDelegate: () => void;

    constructor(model: IBugBash) {
        this._model = {...model};
        this._model.configTemplates = {...model.configTemplates};

        this._originalModel = {...model};
        this._originalModel.configTemplates = {...model.configTemplates};
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

    public renew() {
        this._originalModel = {...this._model};
        this.fireChanged();
    }

    public isDirty(): boolean {        
        return !Utils_String.equals(this._model.title, this._originalModel.title)
            || !Utils_String.equals(this._model.workItemType, this._originalModel.workItemType, true)
            || !Utils_String.equals(this._model.description, this._originalModel.description)
            || !Utils_Date.equals(this._model.startTime, this._originalModel.startTime)
            || !Utils_Date.equals(this._model.endTime, this._originalModel.endTime)
            || !Utils_String.equals(this._model.templateId, this._originalModel.templateId, true)
            || !Utils_Array.arrayEquals(this._model.manualFields, this._originalModel.manualFields, (item1: string, item2: string) => Utils_String.equals(item1, item2, true))
            || this._hasConfigTemplateChanged();
    }

    private _hasConfigTemplateChanged(): boolean {
        let keys1 = Object.keys(this._model.configTemplates);
        let keys2 = Object.keys(this._originalModel.configTemplates);

        if (keys1.length !== keys2.length) {
            return true;
        }

        for (let key of keys1) {
            if (!Utils_String.equals(this._model.configTemplates[key], this._originalModel.configTemplates[key], true)) {
                return true;
            }
        }

        return false;
    }

    public isValid(): boolean {
        return this._model.title.trim().length > 0
            && this._model.title.length <= 128
            && this._model.manualFields.length > 0
            && this._model.workItemType.trim().length > 0
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
        this._model.templateId = "";  // reset template
        this._model.configTemplates = {};
        this.fireChanged();
    }

    public updateConfigTemplate(configName: string, templateId: string) {
        this._model.configTemplates[configName] = templateId;
        this.fireChanged();
    }

    public updateDescription(newDescription: string) {
        this._model.description = newDescription;
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

    public updateTemplate(templateId: string) {
        this._model.templateId = templateId;
        this.fireChanged();
    }

    public updateManualFields(fieldNames: string[]) {
        this._model.manualFields = fieldNames;
        this.fireChanged();
    }
}