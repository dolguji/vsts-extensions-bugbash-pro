import { ExtensionDataManager } from "VSTS_Extension/Utilities/ExtensionDataManager";

import { BugBashItemCommentActionsHub } from "./ActionsHub";
import { IBugBashItemComment } from "../Interfaces";
import { StoresHub } from "../Stores/StoresHub";

export module BugBashItemCommentActions {
    export function clearComments() {
        BugBashItemCommentActionsHub.ClearComments.invoke(null);
    }

    export async function initializeComments(bugBashItemId: string) {
        if (StoresHub.bugBashItemCommentStore.isLoaded(bugBashItemId)) {
            BugBashItemCommentActionsHub.InitializeComments.invoke(null);
        }
        else if (!StoresHub.bugBashItemCommentStore.isLoading(bugBashItemId)) {
            StoresHub.bugBashItemCommentStore.setLoading(true, bugBashItemId);

            const comments = await ExtensionDataManager.readDocuments<IBugBashItemComment>(getBugBashItemCollectionKey(bugBashItemId), false);
            for(let comment of comments) {
                translateDates(comment);
            }
            
            BugBashItemCommentActionsHub.InitializeComments.invoke({bugBashItemId: bugBashItemId, comments: comments});
            StoresHub.bugBashItemCommentStore.setLoading(false, bugBashItemId);
        }
    } 

    export async function refreshComments(bugBashItemId: string) {
        if (!StoresHub.bugBashItemCommentStore.isLoading(bugBashItemId)) {
            StoresHub.bugBashItemCommentStore.setLoading(true, bugBashItemId);

            const comments = await ExtensionDataManager.readDocuments<IBugBashItemComment>(getBugBashItemCollectionKey(bugBashItemId), false);
            for(let comment of comments) {
                translateDates(comment);
            }
            
            BugBashItemCommentActionsHub.RefreshComments.invoke({bugBashItemId: bugBashItemId, comments: comments});
            StoresHub.bugBashItemCommentStore.setLoading(false, bugBashItemId);
        }
    }

    export async function createComment(bugBashItemId: string, commentString: string) {
        if (!StoresHub.bugBashItemCommentStore.isLoading(bugBashItemId)) {
            StoresHub.bugBashItemCommentStore.setLoading(true, bugBashItemId);

            try {
                let bugBashItemComment: IBugBashItemComment = {
                    id: `${bugBashItemId}_${Date.now().toString()}`,
                    __etag: 0,
                    createdBy: `${VSS.getWebContext().user.name} <${VSS.getWebContext().user.uniqueName}>`,
                    createdDate: new Date(Date.now()),
                    content: commentString
                };

                const savedComment = await ExtensionDataManager.createDocument<IBugBashItemComment>(getBugBashItemCollectionKey(bugBashItemId), bugBashItemComment, false);
                translateDates(savedComment);
                                
                BugBashItemCommentActionsHub.CreateComment.invoke({bugBashItemId: bugBashItemId, comment: savedComment});
                StoresHub.bugBashItemCommentStore.setLoading(false, bugBashItemId);
            }
            catch (e) {
                StoresHub.bugBashItemCommentStore.setLoading(false, bugBashItemId);
                throw e.message;
            }
        }
    }

    function getBugBashItemCollectionKey(bugBashItemId: string): string {
        return `BugBashItemCollection_${bugBashItemId}`;
    }

    function translateDates(bugBashItemComment: IBugBashItemComment) {
        if (typeof bugBashItemComment.createdDate === "string") {
            if ((bugBashItemComment.createdDate as string).trim() === "") {
                bugBashItemComment.createdDate = undefined;
            }
            else {
                bugBashItemComment.createdDate = new Date(bugBashItemComment.createdDate);
            }
        }
    }
}