import { VersionControlChangeType, ItemContentType, GitPush } from "TFS/VersionControl/Contracts";
import Utils_String = require("VSS/Utils/String");

import { IBugBashItem, IBugBashItemViewModel } from "./Interfaces";
import { StoresHub } from "./Stores/StoresHub";

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

export class BugBashItemHelpers {
    public static getViewModel(bugBashItem: IBugBashItem): IBugBashItemViewModel {
        if (bugBashItem == null) {
            return null;
        }

        return {
            originalBugBashItem: {...bugBashItem},
            updatedBugBashItem: {...bugBashItem},
            newComment: ""
        };
    }

    public static getNewViewModel(bugBashId: string): IBugBashItemViewModel {
        const bugBashItem = this.getNewBugBashItem(bugBashId);
        return {
            originalBugBashItem: bugBashItem,
            updatedBugBashItem: {...bugBashItem},
            newComment: ""
        };
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

    public static isNew(bugBashItem: IBugBashItem): boolean {
        return bugBashItem && (bugBashItem.id == null || bugBashItem.id.trim() === "");
    }
    
    public static isAccepted(bugBashItem: IBugBashItem): boolean {
        return bugBashItem && bugBashItem.workItemId != null && bugBashItem.workItemId > 0;
    }

    public static isDirty(viewModel: IBugBashItemViewModel): boolean {
        if (viewModel == null) {
            return null;
        }

        const updatedBugBashItem = viewModel.updatedBugBashItem;
        const originalBugBashItem = viewModel.originalBugBashItem;
        const newComment = viewModel.newComment;

        return !Utils_String.equals(updatedBugBashItem.title, originalBugBashItem.title)
            || !Utils_String.equals(updatedBugBashItem.teamId, originalBugBashItem.teamId)
            || !Utils_String.equals(updatedBugBashItem.description, originalBugBashItem.description)
            || !Utils_String.equals(updatedBugBashItem.rejectReason, originalBugBashItem.rejectReason)
            || Boolean(updatedBugBashItem.rejected) !== Boolean(originalBugBashItem.rejected)
            || (newComment != null && newComment.trim() !== "");
    }

    public static isValid(viewModel: IBugBashItemViewModel): boolean {
        if (viewModel == null) {
            return null;
        }
        
        const updatedBugBashItem = viewModel.updatedBugBashItem;

        let dataValid = updatedBugBashItem.title.trim().length > 0 
            && updatedBugBashItem.title.trim().length <= 256 
            && updatedBugBashItem.teamId.trim().length > 0
            && (!updatedBugBashItem.rejected || (updatedBugBashItem.rejectReason != null && updatedBugBashItem.rejectReason.trim().length > 0 && updatedBugBashItem.rejectReason.trim().length <= 128));

        if (dataValid) {
            dataValid = dataValid && StoresHub.teamStore.getItem(updatedBugBashItem.teamId) != null;
        }

        return dataValid;
    }
}
