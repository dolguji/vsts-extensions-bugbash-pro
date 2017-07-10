import { StoreFactory } from "VSTS_Extension/Flux/Stores/BaseStore";
import { ExtensionDataManager } from "VSTS_Extension/Utilities/ExtensionDataManager";

import { BugBashItemCommentActionsCreator } from "./ActionsCreator";
import { BugBashItemCommentStore } from "../Stores/BugBashItemCommentStore";
import { IBugBashItemComment } from "../Interfaces";

export module BugBashItemCommentActions {
    var bugBashItemCommentStore: BugBashItemCommentStore = StoreFactory.getInstance<BugBashItemCommentStore>(BugBashItemCommentStore);

    export async function initializeComments(bugBashItemId: string) {
        if (bugBashItemCommentStore.isLoaded(bugBashItemId)) {
            BugBashItemCommentActionsCreator.InitializeComments.invoke(null);
        }
        else if (!bugBashItemCommentStore.isLoading(bugBashItemId)) {
            bugBashItemCommentStore.setLoading(true, bugBashItemId);

            try {
                const comments = await ExtensionDataManager.readDocuments<IBugBashItemComment>(getBugBashItemCollectionKey(bugBashItemId), false);
                for(let comment of comments) {
                    translateDates(comment);
                }
                bugBashItemCommentStore.setLoading(false, bugBashItemId);
                bugBashItemCommentStore.setError(null, bugBashItemId);

                BugBashItemCommentActionsCreator.InitializeComments.invoke({bugBashItemId: bugBashItemId, comments: comments});
            }
            catch (e) {
                bugBashItemCommentStore.setLoading(false, bugBashItemId);
                bugBashItemCommentStore.setError(e.message || e, bugBashItemId);
            }
        }
    } 

    export async function refreshComments(bugBashItemId: string) {
        if (!bugBashItemCommentStore.isLoading(bugBashItemId)) {
            bugBashItemCommentStore.setLoading(true, bugBashItemId);

            try {
                const comments = await ExtensionDataManager.readDocuments<IBugBashItemComment>(getBugBashItemCollectionKey(bugBashItemId), false);
                for(let comment of comments) {
                    translateDates(comment);
                }
                bugBashItemCommentStore.setLoading(false, bugBashItemId);
                bugBashItemCommentStore.setError(null, bugBashItemId);

                BugBashItemCommentActionsCreator.RefreshComments.invoke({bugBashItemId: bugBashItemId, comments: comments});
            }
            catch (e) {
                bugBashItemCommentStore.setLoading(false, bugBashItemId);
                bugBashItemCommentStore.setError(e.message || e, bugBashItemId);
            }
        }
    }

    export async function createComment(bugBashItemId: string, comment: string) {
        if (!bugBashItemCommentStore.isLoading(bugBashItemId)) {
            bugBashItemCommentStore.setLoading(true, bugBashItemId);

            try {
                let commentModel: IBugBashItemComment = {
                    id: `${bugBashItemId}_${Date.now().toString()}`,
                    __etag: 0,
                    createdBy: `${VSS.getWebContext().user.name} <${VSS.getWebContext().user.uniqueName}>`,
                    createdDate: new Date(Date.now()),
                    content: comment
                };

                const savedComment = await ExtensionDataManager.createDocument<IBugBashItemComment>(getBugBashItemCollectionKey(bugBashItemId), commentModel, false);
                translateDates(savedComment);
                
                bugBashItemCommentStore.setLoading(false, bugBashItemId);
                bugBashItemCommentStore.setError(null, bugBashItemId);

                BugBashItemCommentActionsCreator.CreateComment.invoke({bugBashItemId: bugBashItemId, comment: savedComment});
            }
            catch (e) {
                bugBashItemCommentStore.setLoading(false, bugBashItemId);
                bugBashItemCommentStore.setError(e.message || e, bugBashItemId);
            }
        }
    }

    function getBugBashItemCollectionKey(itemId: string): string {
        return `BugBashItemCollection_${itemId}`;
    }

    function translateDates(model: IBugBashItemComment) {
        if (typeof model.createdDate === "string") {
            if ((model.createdDate as string).trim() === "") {
                model.createdDate = undefined;
            }
            else {
                model.createdDate = new Date(model.createdDate);
            }
        }
    }
}