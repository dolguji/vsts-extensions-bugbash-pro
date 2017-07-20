import { BaseStore } from "VSTS_Extension/Flux/Stores/BaseStore";
import { IBugBashItemComment } from "../Interfaces";
import { BugBashItemCommentActionsHub } from "../Actions/ActionsHub";

export class BugBashItemCommentStore extends BaseStore<IDictionaryStringTo<IBugBashItemComment[]>, IBugBashItemComment[], string> {
    constructor() {
        super();
        this.items = {};    
    }

    public getItem(bugBashItemId: string): IBugBashItemComment[] {
         return this.items[bugBashItemId.toLowerCase()] || null;
    }

    protected initializeActionListeners() {
        BugBashItemCommentActionsHub.InitializeComments.addListener((commentItems: {bugBashItemId: string, comments: IBugBashItemComment[]}) => {
            if (commentItems) {
                this.items[commentItems.bugBashItemId.toLowerCase()] = commentItems.comments;
            }

            this.emitChanged();
        });

        BugBashItemCommentActionsHub.RefreshComments.addListener((commentItems: {bugBashItemId: string, comments: IBugBashItemComment[]}) => {
            if (commentItems) {
                this.items[commentItems.bugBashItemId.toLowerCase()] = commentItems.comments;
            }

            this.emitChanged();
        });

        BugBashItemCommentActionsHub.CreateComment.addListener((commentItem: {bugBashItemId: string, comment: IBugBashItemComment}) => {
            if (commentItem) {
                if (this.items[commentItem.bugBashItemId.toLowerCase()] == null) {
                    this.items[commentItem.bugBashItemId.toLowerCase()] = [];
                }
                this.items[commentItem.bugBashItemId.toLowerCase()].push(commentItem.comment);
            }

            this.emitChanged();
        });
    } 

    public getKey(): string {
        return "BugBashItemCommentStore";
    }

    protected convertItemKeyToString(key: string): string {
        return key;
    }
}