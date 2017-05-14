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

export interface IBugBashItem {
    id: string;
    bugBashId: string;
    readonly __etag: number;
    title: string;
    comments: IComment[];
    workItemId: number;
    description?: string;    
    createdDate: Date;
    createdBy: string;
    acceptedDate: Date;
    acceptedBy: string;
}

export interface IComment {
    text: string;
    addedBy: string;
    addedDate: Date;
}

export interface IBugBashItemModel {
    model: IBugBashItem;
    originalModel: IBugBashItem;
    newComment?: string;
}