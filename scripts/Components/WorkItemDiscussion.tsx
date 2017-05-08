import * as React from "react";
import { autobind } from "OfficeFabric/Utilities";
import { TextField } from "OfficeFabric/TextField";
import { Label } from "OfficeFabric/Label";
import { CommandBar } from "OfficeFabric/CommandBar";
import { IContextualMenuItem } from "OfficeFabric/components/ContextualMenu/ContextualMenu.Props";

import { WorkItem, WorkItemComment } from "TFS/WorkItemTracking/Contracts";
import * as WitClient from "TFS/WorkItemTracking/RestClient";
import Utils_Date = require("VSS/Utils/Date");


export interface IWorkItemDiscussionProps  {
    workItem: WorkItem;
    onClose: () => void;
}

export interface IWorkItemDiscussionState {
    loading: boolean;
    comments: WorkItemComment[];
    newComment: string;
    error?: string;
}

export class WorkItemDiscussion extends React.Component<IWorkItemDiscussionProps, IWorkItemDiscussionState> {
    
}