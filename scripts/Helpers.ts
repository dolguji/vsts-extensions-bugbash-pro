import { JsonPatchDocument, JsonPatchOperation, Operation } from "VSS/WebApi/Contracts";
import * as WitClient from "TFS/WorkItemTracking/RestClient";
import { WorkItem } from "TFS/WorkItemTracking/Contracts";
import Utils_String = require("VSS/Utils/String");
import { VersionControlChangeType, ItemContentType, GitPush } from "TFS/VersionControl/Contracts";

import { StoreFactory } from "VSTS_Extension/Flux/Stores/BaseStore";
import { TeamStore } from "VSTS_Extension/Flux/Stores/TeamStore";

import { IBugBash, IBugBashItem, IBugBashItemViewModel } from "./Interfaces";

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

export class BugBashHelpers {
    public static getNewBugBash(): IBugBash {
        return {
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
        };
    }
}

export class BugBashItemHelpers {
    public static getNewBugBashItemViewModel(bugBashId: string): IBugBashItemViewModel {        
        return {
            bugBashItem: this.getNewBugBashItem(bugBashId),
            originalBugBashItem: this.getNewBugBashItem(bugBashId),
            newComment: ""
        }
    }

    public static getNewBugBashItem(bugBashId: string): IBugBashItem {
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

    public static getBugBashItemViewModel(bugBashItem: IBugBashItem): IBugBashItemViewModel {
        if (!bugBashItem) {
            return null;
        }

        return {
            bugBashItem: {...bugBashItem},
            originalBugBashItem: {...bugBashItem},
            newComment: ""
        }
    }

    public static isNew(bugBashItem: IBugBashItem): boolean {
        return !bugBashItem.id;
    }

    public static isDirty(bugBashItemViewModel: IBugBashItemViewModel): boolean {        
        let isDirty = !Utils_String.equals(bugBashItemViewModel.bugBashItem.title, bugBashItemViewModel.originalBugBashItem.title)
            || !Utils_String.equals(bugBashItemViewModel.bugBashItem.teamId, bugBashItemViewModel.originalBugBashItem.teamId)
            || !Utils_String.equals(bugBashItemViewModel.bugBashItem.description, bugBashItemViewModel.originalBugBashItem.description)
            || !Utils_String.equals(bugBashItemViewModel.bugBashItem.rejectReason, bugBashItemViewModel.originalBugBashItem.rejectReason)
            || Boolean(bugBashItemViewModel.bugBashItem.rejected) !== Boolean(bugBashItemViewModel.originalBugBashItem.rejected)
            || (bugBashItemViewModel.newComment != null && bugBashItemViewModel.newComment.trim() !== "");

        return isDirty;
    }

    public static isValid(bugBashItem: IBugBashItem): boolean {
        let dataValid = bugBashItem.title.trim().length > 0 
            && bugBashItem.title.trim().length <= 256 
            && bugBashItem.teamId.trim().length > 0
            && (!bugBashItem.rejected || (bugBashItem.rejectReason != null && bugBashItem.rejectReason.trim().length > 0 && bugBashItem.rejectReason.trim().length <= 128));

        if (dataValid) {
            dataValid = dataValid && StoreFactory.getInstance<TeamStore>(TeamStore).getItem(bugBashItem.teamId) != null;
        }

        return dataValid;
    }

    public static isAccepted(bugBashItem: IBugBashItem): boolean {
        return bugBashItem.workItemId != null && bugBashItem.workItemId > 0;
    }
}
