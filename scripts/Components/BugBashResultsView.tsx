import * as React from "react";

import { MessagePanel, MessageType } from "VSTS_Extension/Components/Common/MessagePanel";
import { Loading } from "VSTS_Extension/Components/Common/Loading";
import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { BaseStore } from "VSTS_Extension/Flux/Stores/BaseStore";
import { WorkItemFieldActions } from "VSTS_Extension/Flux/Actions/WorkItemFieldActions";
import { TeamActions } from "VSTS_Extension/Flux/Actions/TeamActions";

import { Overlay } from "OfficeFabric/Overlay";

import Utils_String = require("VSS/Utils/String");

import { IBugBash } from "../Interfaces";
import { BugBashResults } from "./BugBashResults";
import { StoresHub } from "../Stores/StoresHub";
import { BugBashActions } from "../Actions/BugBashActions";

interface IBugBashResultsViewState extends IBaseComponentState {
    bugBash: IBugBash;
    loading?: boolean;
}

interface IBugBashResultsViewProps extends IBaseComponentProps {
    id: string;
}

export class BugBashResultsView extends BaseComponent<IBugBashResultsViewProps, IBugBashResultsViewState> {
    protected initializeState() {
        this.state = {
            bugBash: null,
            loading: true
        };
    }

    protected getStores(): BaseStore<any, any, any>[] {
        return [StoresHub.bugBashStore, StoresHub.workItemFieldStore, StoresHub.teamStore];
    }

    protected getStoresState(): IBugBashResultsViewState {
        if (this.props.id) {
            const bugBash = StoresHub.bugBashStore.getItem(this.props.id);
            const loading = StoresHub.workItemFieldStore.isLoading() || StoresHub.teamStore.isLoading() || StoresHub.bugBashStore.isLoading(this.props.id);
            let newState = {
                loading: loading
            } as IBugBashResultsViewState;
            
            if (!loading) {
                newState = {...newState, bugBash: bugBash && {...bugBash}};
            }

            return newState;
        }
        else {
            return {
                loading: StoresHub.workItemFieldStore.isLoading() || StoresHub.teamStore.isLoading() || StoresHub.bugBashStore.isLoading()
            } as IBugBashResultsViewState;
        }
    }

    public componentDidMount(): void {
        super.componentDidMount();
        WorkItemFieldActions.initializeWorkItemFields();
        TeamActions.initializeTeams();
        
        this._initializeBugBashItem(this.props.id);
    }  

    public componentWillReceiveProps(nextProps: Readonly<IBugBashResultsViewProps>): void {
        if (nextProps.id !== this.props.id) {
            const bugBash: IBugBash = StoresHub.bugBashStore.getItem(nextProps.id);

            this.updateState({
                bugBash: {...bugBash},
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
        if (!this.state.loading && !this.state.bugBash) {
            return <MessagePanel messageType={MessageType.Error} message="This instance of bug bash doesn't exist." />;
        }
        else if(!this.state.loading && this.state.bugBash && !Utils_String.equals(VSS.getWebContext().project.id, this.state.bugBash.projectId, true)) {
            return <MessagePanel messageType={MessageType.Error} message="This instance of bug bash is out of scope of current project." />;
        }
        else {
            return <div>
                { this.state.loading && <Overlay className="loading-overlay"><Loading /></Overlay> }
                { this.state.bugBash && StoresHub.teamStore.isLoaded() && StoresHub.workItemFieldStore.isLoaded() && <BugBashResults bugBash={this.state.bugBash} /> }
            </div>;
        }            
    }
}