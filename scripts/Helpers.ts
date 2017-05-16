import {JsonPatchDocument, JsonPatchOperation, Operation} from "VSS/WebApi/Contracts";
import * as WitClient from "TFS/WorkItemTracking/RestClient";
import { WorkItem } from "TFS/WorkItemTracking/Contracts";
import Utils_String = require("VSS/Utils/String");

import { IBugBashItem, IBugBashItemViewModel } from "./Interfaces";

export async function confirmAction(condition: boolean, msg: string): Promise<boolean> {
    if (condition) {
        let dialogService: IHostDialogService = await VSS.getService(VSS.ServiceIds.Dialog) as IHostDialogService;
        try {
            await dialogService.openMessageDialog(msg, { useBowtieStyle: true });
            return true;
        }
        catch (e) {
            // user selected "No"" in dialog
            return false;
        }
    }

    return true;
}

export async function createWorkItem(workItemType: string, fieldValues: IDictionaryStringTo<string>): Promise<WorkItem> {
    let patchDocument: JsonPatchDocument & JsonPatchOperation[] = [];
    for (let fieldRefName in fieldValues) {
        patchDocument.push({
            op: Operation.Add,
            path: `/fields/${fieldRefName}`,
            value: fieldValues[fieldRefName]
        } as JsonPatchOperation);
    }

    return await WitClient.getClient().createWorkItem(patchDocument, VSS.getWebContext().project.id, workItemType);
}

export class BugBashItemHelpers {
    public static getNewItemViewModel(bugBashId: string): IBugBashItemViewModel {        
        return {
            model: this.getNewItem(bugBashId),
            originalModel: this.getNewItem(bugBashId)
        }
    }

    public static getNewItem(bugBashId: string): IBugBashItem {
        return {
            id: "",
            bugBashId: bugBashId,
            __etag: 0,
            title: "",
            description: "",
            areaPath: "",
            workItemId: 0,
            createdDate: null,
            createdBy: ""
        };
    }

    public static deepCopy(model: IBugBashItem): IBugBashItem {
        return {
            id: model.id,
            bugBashId: model.bugBashId,
            __etag: model.__etag,
            title: model.title,
            areaPath: "",
            description: model.description,
            workItemId: model.workItemId,
            createdDate: model.createdDate,
            createdBy: model.createdBy
        }    
    }

    public static getItemViewModel(model: IBugBashItem): IBugBashItemViewModel {
        return {
            model: this.deepCopy(model),
            originalModel: this.deepCopy(model)
        }
    }

    public static isNew(model: IBugBashItem): boolean {
        return !model.id;
    }

    public static isDirty(viewModel: IBugBashItemViewModel): boolean {        
        return !Utils_String.equals(viewModel.model.title, viewModel.originalModel.title)
            || !Utils_String.equals(viewModel.model.description, viewModel.originalModel.description);
    }

    public static isValid(model: IBugBashItem): boolean {
        return model.title.trim().length > 0 && model.title.trim().length <= 256;
    }

    public static isAccepted(model: IBugBashItem): boolean {
        return model.workItemId != null && model.workItemId > 0;
    }
}
