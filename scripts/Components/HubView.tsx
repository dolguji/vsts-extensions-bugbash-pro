import * as React from "react";

import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { StoreFactory } from "VSTS_Extension/Stores/BaseStore"
import { WorkItemFieldStore } from "VSTS_Extension/Stores/WorkItemFieldStore"
import { WorkItemTypeStore } from "VSTS_Extension/Stores/WorkItemTypeStore"
import { WorkItemTemplateStore } from "VSTS_Extension/Stores/WorkItemTemplateStore";

export interface IHubViewProps extends IBaseComponentProps {
    id?: string;
}

export interface IHubViewState extends IBaseComponentState {
    loading: boolean;
}

export abstract class HubView<T extends IHubViewState> extends BaseComponent<IBaseComponentProps, T> {
    constructor(props: IHubViewProps, context: any) {
        super(props, context);

        this.state = this.getStateFromStore();
    }

    public componentDidMount() {
        super.componentDidMount();
        
        this.props.context.stores.bugBashItemStore.addChangedListener(this._onStoreChanged);
        this.props.context.stores.workItemFieldStore.addChangedListener(this._onStoreChanged);
        this.props.context.stores.workItemTemplateStore.addChangedListener(this._onStoreChanged);
        this.props.context.stores.workItemTypeStore.addChangedListener(this._onStoreChanged);
    }

    public componentWillUnmount() {
        super.componentWillUnmount();
        this.props.context.stores.bugBashItemStore.removeChangedListener(this._onStoreChanged);
        this.props.context.stores.workItemFieldStore.removeChangedListener(this._onStoreChanged);
        this.props.context.stores.workItemTemplateStore.removeChangedListener(this._onStoreChanged);
        this.props.context.stores.workItemTypeStore.removeChangedListener(this._onStoreChanged);
    }

    private _onStoreChanged = (handler: IEventHandler): void => {
        let newState = this.getStateFromStore();
        this.setState(newState);
    }

    protected abstract getStateFromStore(): T;
}