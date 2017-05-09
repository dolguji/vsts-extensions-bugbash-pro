export module UrlActions {
    export var ACTION_NEW = "new";
    export var ACTION_ALL = "all";
    export var ACTION_VIEW = "view";
    export var ACTION_EDIT = "edit";
}

export interface IBugBash {
    id: string;
    readonly __etag: number;
    title: string;
    workItemType: string;
    projectId: string;
    itemDescriptionField: string;
    autoAccept: boolean;
    description?: string;    
    startTime?: Date;
    endTime?: Date;
    acceptTemplate?: {team: string, templateId: string};
}

// bug bash item document for 1 bug bash collection
export interface IBugBashItemDocument {
    id: string;
    bugBashId: string;
    readonly __etag: number;
    title: string;    
    description?: string;
    workItemId?: number;
    createdDate: Date;
    createdBy: string;
}

// work item comment document for 1 work item collection
export interface ICommentDocument {
    id: string;
    bugBashItemDocumentId: string;
    readonly __etag: number;
    comment: string;
    addedDate: Date;
    addedBy: string;
}