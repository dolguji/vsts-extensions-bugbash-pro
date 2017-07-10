import { JsonPatchDocument, JsonPatchOperation, Operation } from "VSS/WebApi/Contracts";
import * as WitClient from "TFS/WorkItemTracking/RestClient";
import { WorkItem } from "TFS/WorkItemTracking/Contracts";
import Utils_String = require("VSS/Utils/String");
import { VersionControlChangeType, ItemContentType, GitPush } from "TFS/VersionControl/Contracts";

import { StoreFactory } from "VSTS_Extension/Flux/Stores/BaseStore";
import { TeamStore } from "VSTS_Extension/Flux/Stores/TeamStore";

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

export function buildGitPush(path: string, oldObjectId: string, changeType: VersionControlChangeType, content: string, contentType: ItemContentType): GitPush {
    const commits = [{
    comment: "Adding new image from bug bash pro extension",
    changes: [
        {
        changeType,
        item: {path},
        newContent: content !== undefined ? {
            content,
            contentType,
        } : undefined,
        }],
    }];

    return {
        refUpdates: [{
            name: "refs/heads/master",
            oldObjectId: oldObjectId,
        }],
        commits,
    } as GitPush;
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

export async function updateWorkItem(workItemId: number, fieldValues: IDictionaryStringTo<string>): Promise<WorkItem> {
    let patchDocument: JsonPatchDocument & JsonPatchOperation[] = [];
    for (let fieldRefName in fieldValues) {
        patchDocument.push({
            op: Operation.Add,
            path: `/fields/${fieldRefName}`,
            value: fieldValues[fieldRefName]
        } as JsonPatchOperation);
    }

    return await WitClient.getClient().updateWorkItem(patchDocument, workItemId);
}

export class BugBashItemHelpers {
    public static getNewItemViewModel(bugBashId: string): IBugBashItemViewModel {        
        return {
            model: this.getNewItem(bugBashId),
            originalModel: this.getNewItem(bugBashId),
            newComment: ""
        }
    }

    public static getNewItem(bugBashId: string): IBugBashItem {
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

    public static deepCopy(model: IBugBashItem): IBugBashItem {
        return {
            id: model.id,
            bugBashId: model.bugBashId,
            __etag: model.__etag,
            title: model.title,
            teamId: model.teamId,
            description: model.description,
            workItemId: model.workItemId,
            createdDate: model.createdDate,
            createdBy: model.createdBy,
            rejected: model.rejected,
            rejectReason: model.rejectReason,
            rejectedBy: model.rejectedBy
        }    
    }

    public static getItemViewModel(model: IBugBashItem): IBugBashItemViewModel {
        return {
            model: this.deepCopy(model),
            originalModel: this.deepCopy(model),
            newComment: ""
        }
    }

    public static isNew(model: IBugBashItem): boolean {
        return !model.id;
    }

    public static isDirty(viewModel: IBugBashItemViewModel): boolean {        
        let isDirty = !Utils_String.equals(viewModel.model.title, viewModel.originalModel.title)
            || !Utils_String.equals(viewModel.model.teamId, viewModel.originalModel.teamId)
            || !Utils_String.equals(viewModel.model.description, viewModel.originalModel.description)
            || !Utils_String.equals(viewModel.model.rejectReason, viewModel.originalModel.rejectReason)
            || Boolean(viewModel.model.rejected) !== Boolean(viewModel.originalModel.rejected)
            || (viewModel.newComment != null && viewModel.newComment.trim() !== "");

        return isDirty;
    }

    public static isValid(model: IBugBashItem): boolean {
        let dataValid = model.title.trim().length > 0 
            && model.title.trim().length <= 256 
            && model.teamId.trim().length > 0
            && (!model.rejected || (model.rejectReason != null && model.rejectReason.trim().length > 0 && model.rejectReason.trim().length <= 128));

        if (dataValid) {
            dataValid = dataValid && StoreFactory.getInstance<TeamStore>(TeamStore).getItem(model.teamId) != null;
        }

        return dataValid;
    }

    public static isAccepted(model: IBugBashItem): boolean {
        return model.workItemId != null && model.workItemId > 0;
    }
}
