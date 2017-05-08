import "../css/App.scss";

import * as React from "react";
import * as ReactDOM from "react-dom";

import { Fabric } from "OfficeFabric/Fabric";
import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { LazyLoad } from "VSTS_Extension/Components/Common/LazyLoad";
import { Loading } from "VSTS_Extension/Components/Common/Loading";

import { HostNavigationService } from "VSS/SDK/Services/Navigation";

import { UrlActions } from "./Models";

export enum HubViewMode {
    All,
    New,
    View,
    Edit,
    Loading
}

export interface IHubState extends IBaseComponentState {
    hubViewMode?: HubViewMode;
    id?: string;
}

export class Hub extends BaseComponent<IBaseComponentProps, IHubState> {
    protected initializeState(): void {
        this.state = {};
    }

    protected initialize() { 
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
                    view = <LazyLoad module="scripts/AllBugBashesView">
                            {(AllBugBashesView) => (
                                <AllBugBashesView.AllBugBashesView />
                            )}
                        </LazyLoad>;
                    break;
                case HubViewMode.New:
                    view = <LazyLoad module="scripts/NewBugBashView">
                            {(NewBugBashView) => (
                                <NewBugBashView.NewBugBashView />
                            )}
                        </LazyLoad>;
                    break;
                case HubViewMode.Edit:
                    view = <LazyLoad module="scripts/EditBugBashView">
                            {(EditBugBashView) => (
                                <EditBugBashView.EditBugBashView id={this.state.id} />
                            )}
                        </LazyLoad>;
                    break;
                case HubViewMode.View:
                    view = <LazyLoad module="scripts/ViewBugBashView">
                            {(ViewBugBashView) => (
                                <ViewBugBashView.ViewBugBashView id={this.state.id} />
                            )}
                        </LazyLoad>;
                    break;
                default:
                    view = <Loading />;
                    break;
            }
        }

        return (
            <Fabric className="fabric-container">
                {view}
            </Fabric>
        );
    }

    private async _initialize() {
        this.updateState({ hubViewMode: HubViewMode.Loading });

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
            this.updateState({ hubViewMode: HubViewMode.All });
        }, true);

        navigationService.attachNavigate(UrlActions.ACTION_NEW, () => {
            this.updateState({ hubViewMode: HubViewMode.New });
        }, true);

        navigationService.attachNavigate(UrlActions.ACTION_EDIT, async () => {
            const state = await navigationService.getCurrentState();
            if (state.id) {
                this.updateState({ hubViewMode: HubViewMode.Edit, id: state.id });
            }
        }, true);

        navigationService.attachNavigate(UrlActions.ACTION_VIEW, async () => {
            const state = await navigationService.getCurrentState();
            if (state.id) {
                this.updateState({ hubViewMode: HubViewMode.View, id: state.id });
            }
        }, true);
    }
}

export function init() {
    ReactDOM.render(<Hub />, $("#ext-container")[0]);
}