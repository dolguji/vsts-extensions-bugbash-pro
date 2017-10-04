export namespace UrlActions {
    export const ACTION_ALL = "all";
    export const ACTION_RESULTS = "results";
    export const ACTION_EDIT = "edit";
    export const ACTION_CHARTS = "charts";
    export const ACTION_DETAILS = "details";
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
    export const BugBashDetailsError = "BugBashDetailsError";
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

export namespace SizeLimits {
    export const TitleFieldMaxLength = 255;
    export const RejectFieldMaxLength = 128;
}

export enum DirectoryPagePivotKeys {
    Ongoing = "ongoing",
    Upcoming = "upcoming",
    Past = "past"
}

export enum BugBashViewPivotKeys {
    Results = "results",
    Edit = "edit",
    Charts = "charts",
    Details = "details"
}