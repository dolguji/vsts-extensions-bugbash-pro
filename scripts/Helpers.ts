import { VersionControlChangeType, ItemContentType, GitPush } from "TFS/VersionControl/Contracts";
import * as GitClient from "TFS/VersionControl/GitRestClient";
import * as Context from "VSS/Context";
import { HostNavigationService } from "VSS/SDK/Services/Navigation";

import { StoresHub } from "./Stores/StoresHub";

export async function navigate(action: string, data?: IDictionaryStringTo<any>, replaceHistoryEntry?: boolean, mergeWithCurrentState?: boolean, windowTitle?: string, suppressNavigate?: boolean) {
    let navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
    navigationService.updateHistoryEntry(action, data, replaceHistoryEntry, mergeWithCurrentState, windowTitle, suppressNavigate);
}

export function getBugBashUrl(bugBashId: string, viewName?: string): string {
    const pageContext = Context.getPageContext();
    const navigation = pageContext.navigation;
    const webContext = VSS.getWebContext();
    let url = `${webContext.collection.uri}/${webContext.project.name}/_${navigation.currentController}/${navigation.currentAction}/${navigation.currentParameters}?_a=${viewName}`;

    if (bugBashId) {
        url += `&id=${bugBashId}`;
    }

    return url;
}

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

export async function copyImageToGitRepo(imageData: any, gitFolderSuffix: string): Promise<string> {
    const settings = StoresHub.bugBashSettingsStore.getAll();
    if (settings && settings.gitMediaRepo) {
        const dataStartIndex = imageData.indexOf(",") + 1;
        const metaPart = imageData.substring(5, dataStartIndex - 1);
        const dataPart = imageData.substring(dataStartIndex);

        const extension = metaPart.split(";")[0].split("/").pop();
        const fileName = `pastedImage_${Date.now().toString()}.${extension}`;
        const gitPath = `BugBash_${gitFolderSuffix}/pastedImages/${fileName}`;
        const projectId = VSS.getWebContext().project.id;

        try {
            const gitClient = GitClient.getClient();
            const gitItem = await gitClient.getItem(settings.gitMediaRepo, "/", projectId);
            const pushModel = buildGitPush(gitPath, gitItem.commitId, VersionControlChangeType.Add, dataPart, ItemContentType.Base64Encoded);
            await gitClient.createPush(pushModel, settings.gitMediaRepo, projectId);

            return `${VSS.getWebContext().collection.uri}/${VSS.getWebContext().project.id}/_api/_versioncontrol/itemContent?repositoryId=${settings.gitMediaRepo}&path=${gitPath}&version=GBmaster&contentOnly=true`;
        }
        catch (e) {
            throw `Image copy failed: ${e.message}`;
        }
    }
    else {
        throw "Image copy failed. No Git repo is setup yet to store image files. Please setup a git repo in Bug Bash settings to store media and attachments.";
    }
}

function buildGitPush(path: string, oldObjectId: string, changeType: VersionControlChangeType, content: string, contentType: ItemContentType): GitPush {
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