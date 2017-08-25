import { IBugBash } from "../Interfaces";
import { BugBashActions } from "../Actions/BugBashActions";

import Utils_String = require("VSS/Utils/String");
import Utils_Date = require("VSS/Utils/Date");

export class BugBash {
    public static getNewBugBashModel(): IBugBash {
        return {
            id: "",
            title: "New Bug Bash",
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
        };
    }

    private _originalModel: IBugBash;
    private _updatedModel: IBugBash;

    get originalModel(): IBugBash {
        return {...this._originalModel};
    }
    
    get updatedModel(): IBugBash {
        return {...this._updatedModel};
    }

    get id(): string {
        return this._originalModel.id;
    }

    get projectId(): string {
        return this._originalModel.projectId;
    }

    get version(): number {
        return this._originalModel.__etag;
    }

    constructor(model?: IBugBash) {
        const bugBashModel = model || BugBash.getNewBugBashModel();
        this._originalModel = {...bugBashModel};
        this._updatedModel = {...bugBashModel};
    }

    public update(model: IBugBash, fireChange: boolean = true) {
        this._updatedModel = {...model};

        if (fireChange) {
            BugBashActions.fireStoreChange();
        }        
    }

    public async save() {
        if (this.isDirty() && this.isValid()) {
            if (this.isNew()) {
                BugBashActions.createBugBash(this._updatedModel);
            }
            else {
                BugBashActions.updateBugBash(this._updatedModel);
            }
        }
    }

    public reset(fireChange: boolean = true) {
        if (this.isDirty()) {
            this.update(this._originalModel, fireChange);
        }
    }

    public async refresh() {
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
        const updatedModel = this._updatedModel;
        const originalModel = this._originalModel;

        return !Utils_String.equals(updatedModel.title, originalModel.title)
            || !Utils_String.equals(updatedModel.workItemType, originalModel.workItemType, true)
            || !Utils_String.equals(updatedModel.description, originalModel.description)
            || !Utils_Date.equals(updatedModel.startTime, originalModel.startTime)
            || !Utils_Date.equals(updatedModel.endTime, originalModel.endTime)
            || !Utils_String.equals(updatedModel.itemDescriptionField, originalModel.itemDescriptionField, true)
            || updatedModel.autoAccept !== originalModel.autoAccept
            || !Utils_String.equals(updatedModel.acceptTemplate.team, originalModel.acceptTemplate.team)
            || !Utils_String.equals(updatedModel.acceptTemplate.templateId, originalModel.acceptTemplate.templateId);

    }

    public isValid(): boolean {
        const updatedModel = this._updatedModel;

        return updatedModel.title.trim().length > 0
            && updatedModel.title.length <= 256
            && updatedModel.workItemType.trim().length > 0
            && updatedModel.itemDescriptionField.trim().length > 0
            && (!updatedModel.startTime || !updatedModel.endTime || Utils_Date.defaultComparer(updatedModel.startTime, updatedModel.endTime) < 0);
    }
}