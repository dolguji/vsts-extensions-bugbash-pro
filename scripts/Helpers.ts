import { VersionControlChangeType, ItemContentType, GitPush } from "TFS/VersionControl/Contracts";
import * as GitClient from "TFS/VersionControl/GitRestClient";
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

export async function copyImageToGitRepo(imageData: any, gitFolderSuffix: string, callback) {
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

            const imageUrl = `${VSS.getWebContext().collection.uri}/${VSS.getWebContext().project.id}/_api/_versioncontrol/itemContent?repositoryId=${settings.gitMediaRepo}&path=${gitPath}&version=GBmaster&contentOnly=true`;
            callback(imageUrl);
        }
        catch (e) {
            callback(null);
        }
    }
    else {
        callback(null);
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