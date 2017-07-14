import * as React from "react";

import { MessagePanel, MessageType } from "VSTS_Extension/Components/Common/MessagePanel";
import { Loading } from "VSTS_Extension/Components/Common/Loading";
import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { BaseStore } from "VSTS_Extension/Flux/Stores/BaseStore";
import { WorkItemFieldActions } from "VSTS_Extension/Flux/Actions/WorkItemFieldActions";
import { WorkItemTypeActions } from "VSTS_Extension/Flux/Actions/WorkItemTypeActions";
import { WorkItemTemplateActions } from "VSTS_Extension/Flux/Actions/WorkItemTemplateActions";

import { Overlay } from "OfficeFabric/Overlay";

import Utils_String = require("VSS/Utils/String");

import { IBugBash } from "../Interfaces";
import { BugBashEditor } from "./BugBashEditor";
import { BugBashStore } from "../Stores/BugBashStore";
import { StoresHub } from "../Stores/StoresHub";
import { BugBashHelpers } from "../Helpers";
import { BugBashActions } from "../Actions/BugBashActions";

interface IEditBugBashViewState extends IBaseComponentState {
    model: IBugBash;
    loading?: boolean;
}

interface IEditBugBashViewProps extends IBaseComponentProps {
    id: string;
}

export class EditBugBashView extends BaseComponent<IEditBugBashViewProps, IEditBugBashViewState> {
    protected initializeState() {
        this.state = {
            model: this.props.id ? null : BugBashHelpers.getNewModel(),
            loading: true
        };
    }

    protected getStores(): BaseStore<any, any, any>[] {
        return [StoresHub.bugBashStore, StoresHub.workItemFieldStore, StoresHub.workItemTemplateStore, StoresHub.workItemTypeStore];
    }

    protected getStoresState(): IEditBugBashViewState {
        if (this.props.id) {
            const model = StoresHub.bugBashStore.getItem(this.props.id);
            const loading = StoresHub.workItemTypeStore.isLoading() || StoresHub.workItemFieldStore.isLoading() || StoresHub.workItemTemplateStore.isLoading() || StoresHub.bugBashStore.isLoading(this.props.id);
            let newState = {
                loading: loading
            } as IEditBugBashViewState;
            
            if (!loading) {
                newState = {...newState, model: model && {...model}};
            }

            return newState;
        }
        else {
            return {
                loading: StoresHub.workItemTypeStore.isLoading() || StoresHub.workItemFieldStore.isLoading() || StoresHub.workItemTemplateStore.isLoading() || StoresHub.bugBashStore.isLoading()
            } as IEditBugBashViewState;
        }
    }

    public componentDidMount(): void {
        super.componentDidMount();
        WorkItemFieldActions.initializeWorkItemFields();
        WorkItemTypeActions.initializeWorkItemTypes();
        WorkItemTemplateActions.initializeWorkItemTemplates();
        
        this._initializeBugBashItem(this.props.id);
    }  

    public componentWillReceiveProps(nextProps: Readonly<IEditBugBashViewProps>): void {
        if (nextProps.id !== this.props.id) {
            const model: IBugBash = nextProps.id ? StoresHub.bugBashStore.getItem(nextProps.id) : BugBashHelpers.getNewModel();            

            this.updateState({
                model: {...model},
                loading: false
            });
        }        
    }

    private async _initializeBugBashItem(bugBashId: string) {        
        if (bugBashId) {
            try {
                await BugBashActions.initializeBugBash(bugBashId);
            }
            catch (e) {
                // no-op                
            }
        }
    }

    public render(): JSX.Element {
        if (!this.state.loading && !this.state.model) {
            return <MessagePanel messageType={MessageType.Error} message="This instance of bug bash doesn't exist." />;
        }
        else if(!this.state.loading && this.state.model && !Utils_String.equals(VSS.getWebContext().project.id, this.state.model.projectId, true)) {
            return <MessagePanel messageType={MessageType.Error} message="This instance of bug bash is out of scope of current project." />;
        }
        else {
            return <div>
                { this.state.loading && <Overlay className="loading-overlay"><Loading /></Overlay> }
                { this.state.model && StoresHub.workItemTypeStore.isLoaded() && StoresHub.workItemFieldStore.isLoaded() && StoresHub.workItemTemplateStore.isLoaded() && <BugBashEditor model={this.state.model} /> }
            </div>;
        }            
    }
}