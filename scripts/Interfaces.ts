import { WorkItem } from "TFS/WorkItemTracking/Contracts";

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
    workItemId: number;
    description?: string;    
    createdDate: Date;
    createdBy: string;
    acceptedDate: Date;
    acceptedBy: string;
}

export interface IBugBashItemViewModel {
    model: IBugBashItem;
    originalModel: IBugBashItem;
}

export interface IAcceptedItemViewModel {
    model: IBugBashItem;
    workItem: WorkItem;
}