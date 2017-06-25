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
    teamId: string;
    description?: string;    
    createdDate: Date;
    createdBy: string;
    rejected?: boolean;
    rejectReason?: string;
    rejectedBy?: string;
}

export interface IBugBashItemViewModel {
    model: IBugBashItem;
    originalModel: IBugBashItem;
}

export interface IAcceptedItemViewModel {
    model: IBugBashItem;
    workItem: WorkItem;
}

export interface Settings {
    gitMediaRepo: string;
}