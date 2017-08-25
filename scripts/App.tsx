import "../css/App.scss";

import * as React from "react";
import * as ReactDOM from "react-dom";

import { Fabric } from "OfficeFabric/Fabric";
import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { Loading } from "VSTS_Extension/Components/Common/Loading";

import { HostNavigationService } from "VSS/SDK/Services/Navigation";
import { getAsyncLoadedComponent } from "VSS/Flux/AsyncLoadedComponent";

import { UrlActions } from "./Constants";
import { BugBashView } from "./Components/BugBashView";
import * as AllBugBashesView_Async from "./Components/AllBugBashesView";

export enum HubViewMode {
    All,
    New,
    View,
    Edit,
    Charts
}

export interface IHubState extends IBaseComponentState {
    hubViewMode: HubViewMode;
    bugBashId?: string;
}

const AsyncAllBugBashView = getAsyncLoadedComponent(
    ["scripts/AllBugBashesView"],
    (m: typeof AllBugBashesView_Async) => m.AllBugBashesView,
    () => <Loading />);

export class Hub extends BaseComponent<IBaseComponentProps, IHubState> {
    protected initializeState() {
        this.state = {
            hubViewMode: null
        };
    }

    public componentDidMount() { 
        super.componentDidMount();
        this._initialize();
    }

    public render(): JSX.Element {
        let view;

        if (this.state.hubViewMode == null) {
            view = <Loading />;
        }
        else {
            switch (this.state.hubViewMode) {            
                case HubViewMode.All:
                    view = <AsyncAllBugBashView />;
                    break;
                case HubViewMode.New:
                case HubViewMode.Edit:
                    view = <BugBashView pivotKey={UrlActions.ACTION_EDIT} bugBashId={this.state.bugBashId} />;
                    break;
                case HubViewMode.View:
                    view = <BugBashView pivotKey={UrlActions.ACTION_RESULTS} bugBashId={this.state.bugBashId} />;
                    break;
                case HubViewMode.Charts:
                    view = <BugBashView pivotKey={UrlActions.ACTION_CHARTS} bugBashId={this.state.bugBashId} />;
                    break;
                default:
                    view = <Loading />;
                    break;
            }
        }

        return (
            <Fabric className="fabric-container bowtie-fabric">
                {view}
            </Fabric>
        );
    }

    private async _initialize() {
        this._attachNavigate();

        const navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
        const state = await navigationService.getCurrentState();
        if (!state.action) {
            navigationService.updateHistoryEntry(UrlActions.ACTION_ALL, null, true);
        }        
    }

    private async _attachNavigate() {        
        const navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;

        navigationService.attachNavigate(UrlActions.ACTION_ALL, () => {
            this.updateState({ hubViewMode: HubViewMode.All, bugBashId: null });
        }, true);

        navigationService.attachNavigate(UrlActions.ACTION_EDIT, async () => {
            const state = await navigationService.getCurrentState();
            this.updateState({ hubViewMode: HubViewMode.Edit, bugBashId: state.id || null });
        }, true);

        navigationService.attachNavigate(UrlActions.ACTION_RESULTS, async () => {
            const state = await navigationService.getCurrentState();
            this.updateState({ hubViewMode: HubViewMode.View, bugBashId: state.id || null });
        }, true);

        navigationService.attachNavigate(UrlActions.ACTION_CHARTS, async () => {
            const state = await navigationService.getCurrentState();
            this.updateState({ hubViewMode: HubViewMode.Charts, bugBashId: state.id || null });
        }, true);
    }
}

export function init() {
    ReactDOM.render(<Hub />, $("#ext-container")[0]);
}