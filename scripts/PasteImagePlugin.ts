import * as GitClient from "TFS/VersionControl/GitRestClient";
import { GitRepository, VersionControlChangeType, ItemContentType, GitPush } from "TFS/VersionControl/Contracts";
import { StoresHub } from "./Stores/StoresHub";
import { SettingsStore } from "./Stores/SettingsStore";

(function ($) {
    'use strict';

    $.extend(true, ($ as any).trumbowyg, {
        plugins: {
            pasteImage: {
                init: function (trumbowyg) {
                    trumbowyg.pasteHandlers.push(function (pasteEvent) {
                        try {
                            let items = (pasteEvent.originalEvent || pasteEvent).clipboardData.items, reader;
                            
                            for (var i = items.length -1; i >= 0; i += 1) {
                                if (items[i].type.indexOf("image") === 0) {
                                    reader = new FileReader();
                                    reader.onloadend = async (event) => {
                                        const data = event.target.result;
                                        const settings = StoresHub.settingsStore.getAll();
                                        if (settings && settings.gitMediaRepo) {
                                            const dataStartIndex = data.indexOf(",") + 1;
                                            const metaPart = data.substring(5, dataStartIndex - 1);
                                            const dataPart = data.substring(dataStartIndex);

                                            const extension = metaPart.split(";")[0].split("/").pop();
                                            const fileName = `pastedImage_${Date.now().toString()}.${extension}`;
                                            const gitPath = `/pastedImages/${fileName}`;
                                            const projectId = VSS.getWebContext().project.id;

                                            try {
                                                const gitClient = GitClient.getClient();
                                                const gitItem = await gitClient.getItem(settings.gitMediaRepo, "/", projectId);
                                                const pushModel = _buildGitPush(gitPath, gitItem.commitId, VersionControlChangeType.Add, dataPart, ItemContentType.Base64Encoded);
                                                await gitClient.createPush(pushModel, settings.gitMediaRepo, projectId);

                                                const imageUrl = `${VSS.getWebContext().collection.uri}/${VSS.getWebContext().project.id}/_api/_versioncontrol/itemContent?repositoryId=${settings.gitMediaRepo}&path=${gitPath}&version=GBmaster&contentOnly=true`;
                                                trumbowyg.execCmd('insertImage', imageUrl, undefined, true);
                                            }
                                            catch (e) {
                                                trumbowyg.execCmd('insertImage', data, undefined, true);
                                            }
                                        }
                                        else {
                                            trumbowyg.execCmd('insertImage', data, undefined, true);
                                        }
                                    };
                                    reader.readAsDataURL(items[i].getAsFile());
                                    break;
                                }
                            }
                        } catch (c) {
                        }
                    });
                }
            }
        }
    });
})(jQuery);

function _buildGitPush(path: string, oldObjectId: string, changeType: VersionControlChangeType, content: string, contentType: ItemContentType): GitPush {
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

    return <GitPush>{
      refUpdates: [{
        name: "refs/heads/master",
        oldObjectId: oldObjectId,
      }],
      commits,
    };
}