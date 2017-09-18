import { Action } from "MB/Flux/Actions/Action";

import { ILongText, IUserSettings, IBugBashSettings, IBugBash, IBugBashItemComment, IBugBashItemActionData, IBugBashItemsActionData, IBugBashItemIdActionData } from "../Interfaces";

export namespace BugBashClientActionsHub {
    export const SelectedBugBashItemChanged = new Action<string>();
}

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
    export const InitializeBugBashItems = new Action<IBugBashItemsActionData>();
    export const RefreshBugBashItems = new Action<IBugBashItemsActionData>();
    export const RefreshBugBashItem = new Action<IBugBashItemActionData>();
    export const CreateBugBashItem = new Action<IBugBashItemActionData>();
    export const UpdateBugBashItem = new Action<IBugBashItemActionData>();
    export const DeleteBugBashItem = new Action<IBugBashItemIdActionData>();
    export const AcceptBugBashItem = new Action<IBugBashItemActionData>();
    export const FireStoreChange = new Action<void>();
}

export namespace BugBashErrorMessageActionsHub {
    export const PushErrorMessage = new Action<{errorMessage: string, errorKey: string}>();
    export const DismissErrorMessage = new Action<string>();
}

export namespace LongTextActionsHub {
    export const AddOrUpdateLongText = new Action<ILongText>();
    export const FireStoreChange = new Action<void>();
}