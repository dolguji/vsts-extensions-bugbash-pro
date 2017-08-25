import { BugBashErrorMessageActionsHub } from "./ActionsHub";

export module BugBashErrorMessageActions {
    export function showErrorMessage(errorMessage: string, errorKey: string) {
        BugBashErrorMessageActionsHub.PushErrorMessage.invoke({errorMessage: errorMessage, errorKey: errorKey});
    }

    export function dismissErrorMessage(errorKey: string) {
        return BugBashErrorMessageActionsHub.DismissErrorMessage.invoke(errorKey);
    }
}