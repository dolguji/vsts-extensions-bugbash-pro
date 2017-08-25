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
    acceptTemplateTeam?: string;
    acceptTemplateId?: string;
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

export interface IBugBashItemComment {
    id: string;
    readonly __etag: number;
    content: string;
    createdDate: Date;
    createdBy: string;
}

export interface IBugBashItemViewModel {
    originalBugBashItem: IBugBashItem;
    updatedBugBashItem: IBugBashItem;
    newComment?: string;
}

export interface IBugBashSettings {
    gitMediaRepo: string;
}

export interface IUserSettings {
    id: string;
    readonly __etag: number;
    associatedTeam: string;
}

export interface INameValuePair {
    name: string;
    value: number;
    members?: INameValuePair[];
}