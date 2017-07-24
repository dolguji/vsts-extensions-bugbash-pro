import { VersionControlChangeType, ItemContentType, GitPush } from "TFS/VersionControl/Contracts";
import Utils_String = require("VSS/Utils/String");
import Utils_Date = require("VSS/Utils/Date");

import { IBugBash, IBugBashItem, IBugBashViewModel, IBugBashItemViewModel } from "./Interfaces";
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

export class BugBashHelpers {
    public static getViewModel(bugBash: IBugBash): IBugBashViewModel {
        if (bugBash == null) {
            return null;
        }
        
        return {
            originalBugBash: {...bugBash},
            updatedBugBash: {...bugBash}
        };
    }

    public static getNewViewModel(): IBugBashViewModel {
        const bugBash = this.getNewBugBash();
        return {
            originalBugBash: bugBash,
            updatedBugBash: {...bugBash}
        };
    }

    public static getNewBugBash(): IBugBash {
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

    public static isNew(bugBash: IBugBash): boolean {
        return bugBash && (bugBash.id == null || bugBash.id.trim() === "");
    }

    public static isDirty(viewModel: IBugBashViewModel): boolean {
        if (viewModel == null) {
            return false;
        }

        const updatedBugBash = viewModel.updatedBugBash;
        const originalBugBash = viewModel.originalBugBash;

        return !Utils_String.equals(updatedBugBash.title, originalBugBash.title)
            || !Utils_String.equals(updatedBugBash.workItemType, originalBugBash.workItemType, true)
            || !Utils_String.equals(updatedBugBash.description, originalBugBash.description)
            || !Utils_Date.equals(updatedBugBash.startTime, originalBugBash.startTime)
            || !Utils_Date.equals(updatedBugBash.endTime, originalBugBash.endTime)
            || !Utils_String.equals(updatedBugBash.itemDescriptionField, originalBugBash.itemDescriptionField, true)
            || updatedBugBash.autoAccept !== originalBugBash.autoAccept
            || !Utils_String.equals(updatedBugBash.acceptTemplate.team, originalBugBash.acceptTemplate.team)
            || !Utils_String.equals(updatedBugBash.acceptTemplate.templateId, originalBugBash.acceptTemplate.templateId);

    }

    public static isValid(viewModel: IBugBashViewModel): boolean {
        if (viewModel == null) {
            return false;
        }
        
        const updatedBugBash = viewModel.updatedBugBash;

        return updatedBugBash.title.trim().length > 0
            && updatedBugBash.title.length <= 256
            && updatedBugBash.workItemType.trim().length > 0
            && updatedBugBash.itemDescriptionField.trim().length > 0
            && (!updatedBugBash.startTime || !updatedBugBash.endTime || Utils_Date.defaultComparer(updatedBugBash.startTime, updatedBugBash.endTime) < 0);
    }
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
