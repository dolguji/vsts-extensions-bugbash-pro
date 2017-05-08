import { IBugBashItemDocument } from "../Models";

import Utils_String = require("VSS/Utils/String");

export class BugBashItem {
    public static getNew(bugBashId: string): BugBashItem {
        return new BugBashItem({
            id: "",
            bugBashId: bugBashId,
            title: "",
            __etag: 0,
            description: "",
            workItemId: null
        });
    }

    private _model: IBugBashItemDocument;
    private _originalModel: IBugBashItemDocument;
    private _onChangedDelegate: () => void;

    constructor(model: IBugBashItemDocument) {
        this._model = {...model};
        this._originalModel = {...model};
    }

    public getModel(): IBugBashItemDocument {
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
            || !Utils_String.equals(this._model.description, this._originalModel.description);            
    }

    public isValid(): boolean {
        return this._model.title.trim().length > 0 && this._model.title.length <= 256;            
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

    public updateDescription(newDescription: string) {
        this._model.description = newDescription;
        this.fireChanged();
    }
}