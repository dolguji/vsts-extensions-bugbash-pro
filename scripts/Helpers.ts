import { JsonPatchDocument, JsonPatchOperation, Operation } from "VSS/WebApi/Contracts";
import * as WitClient from "TFS/WorkItemTracking/RestClient";
import { WorkItem } from "TFS/WorkItemTracking/Contracts";
import Utils_String = require("VSS/Utils/String");
import Utils_Date = require("VSS/Utils/Date");
import { VersionControlChangeType, ItemContentType, GitPush } from "TFS/VersionControl/Contracts";

import { IBugBash, IBugBashItem } from "./Interfaces";

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
        return bugBash.id == null || bugBash.id.trim() === "";
    }
}

export class BugBashItemHelpers {    
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
        return bugBashItem.id == null || bugBashItem.id.trim() === "";
    }
    
    public static isAccepted(bugBashItem: IBugBashItem): boolean {
        return bugBashItem.workItemId != null && bugBashItem.workItemId > 0;
    }
}
