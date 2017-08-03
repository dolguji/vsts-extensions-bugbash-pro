import { BugBashErrorMessageActionsHub } from "./ActionsHub";

export module BugBashErrorMessageActions {
    export function showErrorMessage(errorMessage: string): void {
        BugBashErrorMessageActionsHub.PushErrorMessage.invoke(errorMessage);
    }

    export function dismissErrorMessage(): void {
        return BugBashErrorMessageActionsHub.DismissErrorMessage.invoke(null);
    }
}