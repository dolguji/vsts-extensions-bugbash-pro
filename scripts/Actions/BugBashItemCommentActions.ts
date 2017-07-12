import { ExtensionDataManager } from "VSTS_Extension/Utilities/ExtensionDataManager";

import { BugBashItemCommentActionsCreator } from "./ActionsCreator";
import { BugBashItemCommentStore } from "../Stores/BugBashItemCommentStore";
import { IBugBashItemComment } from "../Interfaces";
import { StoresHub } from "../Stores/StoresHub";

export module BugBashItemCommentActions {
    export async function initializeComments(bugBashItemId: string) {
        if (StoresHub.bugBashItemCommentStore.isLoaded(bugBashItemId)) {
            BugBashItemCommentActionsCreator.InitializeComments.invoke(null);
        }
        else if (!StoresHub.bugBashItemCommentStore.isLoading(bugBashItemId)) {
            StoresHub.bugBashItemCommentStore.setLoading(true, bugBashItemId);

            const comments = await ExtensionDataManager.readDocuments<IBugBashItemComment>(getBugBashItemCollectionKey(bugBashItemId), false);
            for(let comment of comments) {
                translateDates(comment);
            }
            
            BugBashItemCommentActionsCreator.InitializeComments.invoke({bugBashItemId: bugBashItemId, comments: comments});
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
            
            BugBashItemCommentActionsCreator.RefreshComments.invoke({bugBashItemId: bugBashItemId, comments: comments});
            StoresHub.bugBashItemCommentStore.setLoading(false, bugBashItemId);
        }
    }

    export async function createComment(bugBashItemId: string, comment: string) {
        if (!StoresHub.bugBashItemCommentStore.isLoading(bugBashItemId)) {
            StoresHub.bugBashItemCommentStore.setLoading(true, bugBashItemId);

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
                                
                BugBashItemCommentActionsCreator.CreateComment.invoke({bugBashItemId: bugBashItemId, comment: savedComment});
                StoresHub.bugBashItemCommentStore.setLoading(false, bugBashItemId);
            }
            catch (e) {
                StoresHub.bugBashItemCommentStore.setLoading(false, bugBashItemId);
                throw e.message;
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