import * as React from "react";
import { ActionsCreator, ActionsHub } from "./Actions/ActionsCreator";
import { StoresHub } from "./Stores/StoresHub";

export interface IHubContext {
    /**
     * Actions hub
     */
    actions: ActionsHub;

    /**
     * Stores hub
     */
    stores: StoresHub;

    /**
     * Action creator
     */
    actionsCreator: ActionsCreator;
}

export module UrlActions {
    export var ACTION_NEW = "new";
    export var ACTION_ALL = "all";
    export var ACTION_VIEW = "view";
    export var ACTION_EDIT = "edit";
}

export module Constants {
    export var BUGBASH_DOCUMENT_COLLECTION_NAME = "bugbashes";
    export var WORKITEM_DOCUMENT_COLLECTION_NAME = "workitems";
    export var COMMENT_DOCUMENT_COLLECTION_NAME = "comments";
    export var ACCEPT_STATUS_CELL_NAME = "BugBashItemAcceptStatus";
    export var ACCEPTED_TEXT = "Accepted";
    export var REJECTED_TEXT = "Rejected";
    export var ACTIONS_CELL_NAME = "Actions";
}

export interface IBugBash {
    id: string,
    title: string;
    workItemType: string;
    readonly __etag: number;
    description?: string;    
    startTime?: Date;
    endTime?: Date;
    projectId: string;
}

// work item document for 1 bug bash collection
export interface IWorkItemDocument {
    id: string,
    title: string;
    readonly __etag: number;
    description?: string;
    attributes?: IDictionaryStringTo<string>;
}

// work item comment document for 1 work item collection
export interface ICommentDocument {
    id: string,
    readonly __etag: number;
    comment: string;
    addedDate: Date;
    addedBy: string;
}

export enum LoadingState {
    Loading,
    Loaded
}

export interface IBaseProps {
    context: IHubContext;
}