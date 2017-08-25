import { Action } from "VSS/Flux/Action";

import { IUserSettings, IBugBashSettings, IBugBash, IBugBashItemComment, IBugBashItem } from "../Interfaces";

export namespace SettingsActionsHub {
    export const InitializeBugBashSettings = new Action<IBugBashSettings>();
    export const UpdateBugBashSettings = new Action<IBugBashSettings>();
    export const InitializeUserSettings = new Action<IUserSettings[]>();
    export const UpdateUserSettings = new Action<IUserSettings>();
}

export namespace BugBashActionsHub {
    export const InitializeAllBugBashes = new Action<IBugBash[]>();
    export const InitializeBugBash = new Action<IBugBash>();
    export const RefreshBugBash = new Action<IBugBash>();
    export const RefreshAllBugBashes = new Action<IBugBash[]>();
    export const UpdateBugBash = new Action<IBugBash>();
    export const DeleteBugBash = new Action<string>();
    export const CreateBugBash = new Action<IBugBash>();
    export const FireStoreChange = new Action<void>();
}

export namespace BugBashItemCommentActionsHub {
    export const InitializeComments = new Action<{bugBashItemId: string, comments: IBugBashItemComment[]}>();
    export const RefreshComments = new Action<{bugBashItemId: string, comments: IBugBashItemComment[]}>();
    export const CreateComment = new Action<{bugBashItemId: string, comment: IBugBashItemComment}>();
    export const ClearComments = new Action<void>();
}

export namespace BugBashItemActionsHub {
    export const InitializeBugBashItems = new Action<{bugBashId: string, bugBashItems: IBugBashItem[]}>();
    export const RefreshBugBashItems = new Action<{bugBashId: string, bugBashItems: IBugBashItem[]}>();
    export const RefreshBugBashItem = new Action<{bugBashId: string, bugBashItem: IBugBashItem}>();
    export const CreateBugBashItem = new Action<{bugBashId: string, bugBashItem: IBugBashItem}>();
    export const UpdateBugBashItem = new Action<{bugBashId: string, bugBashItem: IBugBashItem}>();
    export const DeleteBugBashItem = new Action<{bugBashId: string, bugBashItemId: string}>();
    export const AcceptBugBashItem = new Action<{bugBashId: string, bugBashItem: IBugBashItem}>();
}

export namespace BugBashErrorMessageActionsHub {
    export const PushErrorMessage = new Action<{errorMessage: string, errorKey: string}>();
    export const DismissErrorMessage = new Action<string>();
}