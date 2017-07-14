import { Action } from "VSS/Flux/Action";

import { Settings, IBugBash, IBugBashItemComment, IBugBashItem } from "../Interfaces";

export module SettingsActionsCreator {
    export var InitializeBugBashSettings = new Action<Settings>();
    export var UpdateBugBashSettings = new Action<Settings>();
}

export module BugBashActionsCreator {
    export var InitializeAllBugBashes = new Action<IBugBash[]>();
    export var InitializeBugBash = new Action<IBugBash>();
    export var RefreshBugBash = new Action<IBugBash>();
    export var RefreshAllBugBashes = new Action<IBugBash[]>();
    export var UpdateBugBash = new Action<IBugBash>();
    export var DeleteBugBash = new Action<string>();
    export var CreateBugBash = new Action<IBugBash>();
}

export module BugBashItemCommentActionsCreator {
    export var InitializeComments = new Action<{bugBashItemId: string, comments: IBugBashItemComment[]}>();
    export var RefreshComments = new Action<{bugBashItemId: string, comments: IBugBashItemComment[]}>();
    export var CreateComment = new Action<{bugBashItemId: string, comment: IBugBashItemComment}>();
}

export module BugBashItemActionsCreator {
    export var InitializeItems = new Action<{bugBashId: string, bugBashItems: IBugBashItem[]}>();
    export var RefreshItems = new Action<{bugBashId: string, bugBashItems: IBugBashItem[]}>();
    export var RefreshItem = new Action<{bugBashId: string, bugBashItem: IBugBashItem}>();
    export var CreateItem = new Action<{bugBashId: string, bugBashItem: IBugBashItem}>();
    export var UpdateItem = new Action<{bugBashId: string, bugBashItem: IBugBashItem}>();
    export var AcceptItem = new Action<{bugBashId: string, bugBashItem: IBugBashItem}>();
}