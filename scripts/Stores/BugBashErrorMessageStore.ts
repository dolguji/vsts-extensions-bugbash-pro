import { BaseStore } from "VSTS_Extension/Flux/Stores/BaseStore";
import { BugBashErrorMessageActionsHub } from "../Actions/ActionsHub";

export class BugBashErrorMessageStore extends BaseStore<string, string, void> {
    constructor() {
        super();
        this.items = null;    
    }

    protected initializeActionListeners() {
        BugBashErrorMessageActionsHub.PushErrorMessage.addListener((errorMessage: string) => {
            this.items = errorMessage;
            this.emitChanged();
        });

        BugBashErrorMessageActionsHub.DismissErrorMessage.addListener(() => {
            this.items = null;
            this.emitChanged();
        });
    }   
    
    public getKey(): string {
        return "BugBashErrorMessageStore";
    }

    protected convertItemKeyToString(_key: void): string {
        return "";
    }

    public getItem(_key: void): string {
         return this.items;
    }
}
