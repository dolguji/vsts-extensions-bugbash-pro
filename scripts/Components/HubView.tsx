import * as React from "react";

import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { BaseStore, StoreFactory } from "VSTS_Extension/Stores/BaseStore";
import { WorkItemFieldStore } from "VSTS_Extension/Stores/WorkItemFieldStore";
import { WorkItemTypeStore } from "VSTS_Extension/Stores/WorkItemTypeStore";
import { WorkItemTemplateStore } from "VSTS_Extension/Stores/WorkItemTemplateStore";

import { BugBashItemStore } from "../Stores/BugBashItemStore";

export interface IHubViewProps extends IBaseComponentProps {
    id?: string;
}

export interface IHubViewState extends IBaseComponentState {
    loading: boolean;
}

export abstract class HubView<T extends IHubViewState> extends BaseComponent<IHubViewProps, T> {
    
    protected bugBashItemStore: BugBashItemStore = StoreFactory.getInstance<BugBashItemStore>(BugBashItemStore);
    protected workItemFieldStore: WorkItemFieldStore = StoreFactory.getInstance<WorkItemFieldStore>(WorkItemFieldStore);
    protected workItemTypeStore: WorkItemTypeStore = StoreFactory.getInstance<WorkItemTypeStore>(WorkItemTypeStore);
    protected workItemTemplateStore: WorkItemTemplateStore = StoreFactory.getInstance<WorkItemTemplateStore>(WorkItemTemplateStore);

    protected initializeState(): void {
        this.state = this.getStateFromStore();
    }

    protected getStoresToLoad(): {new (): BaseStore<any, any, any>}[] {
        return [BugBashItemStore, WorkItemFieldStore, WorkItemTemplateStore, WorkItemTypeStore];
    }

    protected onStoreChanged() {
        let newState = this.getStateFromStore();
        this.setState(newState);
    }

    protected abstract getStateFromStore(): T;
}