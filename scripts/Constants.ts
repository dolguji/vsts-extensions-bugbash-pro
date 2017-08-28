export namespace UrlActions {
    export const ACTION_ALL = "all";
    export const ACTION_RESULTS = "results";
    export const ACTION_EDIT = "edit";
    export const ACTION_CHARTS = "charts";
}

export namespace Events {
    export const NewItem = "new-item";
    export const RefreshItems = "refresh-items";
}

export namespace ChartsView {
    export const All = "All Items";
    export const PendingItemsOnly = "Pending Items";
    export const RejectedItemsOnly = "Rejected Items";
    export const AcceptedItemsOnly = "Accepted Items";
}

export namespace ResultsView {
    export const PendingItemsOnly = "Pending Items";
    export const RejectedItemsOnly = "Rejected Items";
    export const AcceptedItemsOnly = "Accepted Items";
}

export namespace ErrorKeys {
    export const DirectoryPageError = "DirectoryPageError";
    export const BugBashError = "BugBashError";
    export const BugBashItemError = "BugBashItemError";
    export const BugBashSettingsError = "BugBashSettingsError";
}

export enum BugBashFieldNames {
    ID = "id",
    Version = "__etag",
    Title = "title",
    WorkItemType = "workItemType",
    ProjectId = "projectId",
    ItemDescriptionField = "itemDescriptionField",
    AutoAccept = "autoAccept",
    Description = "description",
    StartTime = "startTime",
    EndTime = "endTime",
    AcceptTemplateTeam = "acceptTemplateTeam",
    AcceptTemplateId = "acceptTemplateId"
}

export enum BugBashItemFieldNames {
    ID = "id",
    Version = "__etag",
    Title = "title",
    Description = "description",
    BugBashId = "bugBashId",
    WorkItemId = "workItemId",
    TeamId = "teamId",
    CreatedDate = "createdDate",
    CreatedBy = "createdBy",
    Rejected = "rejected",
    RejectReason = "rejectReason",
    RejectedBy = "rejectedBy",
}