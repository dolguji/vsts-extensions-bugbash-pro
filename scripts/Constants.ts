export module UrlActions {
    export var ACTION_ALL = "all";
    export var ACTION_RESULTS = "results";
    export var ACTION_EDIT = "edit";
    export var ACTION_CHARTS = "charts";
}

export module Events {
    export var NewItem = "new-item";
    export var RefreshItems = "refresh-items";
}

export module ChartsView {
    export var All = "All Items";
    export var PendingItemsOnly = "Pending Items";
    export var RejectedItemsOnly = "Rejected Items";
    export var AcceptedItemsOnly = "Accepted Items";
}

export module ResultsView {
    export var PendingItemsOnly = "Pending Items";
    export var RejectedItemsOnly = "Rejected Items";
    export var AcceptedItemsOnly = "Accepted Items";
}

export module ErrorKeys {
    export var DirectoryPageError = "DirectoryPageError";
    export var BugBashError = "BugBashError";
    export var BugBashItemError = "BugBashItemError";
    export var BugBashSettingsError = "BugBashSettingsError";
}
