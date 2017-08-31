import { BaseStore } from "MB/Flux/Stores/BaseStore";
import { BugBashErrorMessageActionsHub } from "../Actions/ActionsHub";

export class BugBashErrorMessageStore extends BaseStore<IDictionaryStringTo<string>, string, string> {
    constructor() {
        super();
        this.items = {};    
    }

    protected initializeActionListeners() {
        BugBashErrorMessageActionsHub.PushErrorMessage.addListener((error: {errorMessage: string, errorKey: string}) => {
            this.items[error.errorKey] = error.errorMessage;
            this.emitChanged();
        });

        BugBashErrorMessageActionsHub.DismissErrorMessage.addListener((errorKey: string) => {
            delete this.items[errorKey];
            this.emitChanged();
        });
    }   
    
    public getKey(): string {
        return "BugBashErrorMessageStore";
    }

    protected convertItemKeyToString(key: string): string {
        return key;
    }

    public getItem(key: string): string {
         return this.items[key];
    }
}