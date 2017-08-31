import "../css/App.scss";

import * as React from "react";
import * as ReactDOM from "react-dom";

import { Fabric } from "OfficeFabric/Fabric";
import { autobind } from "OfficeFabric/Utilities";

import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "MB/Components/BaseComponent";
import { Loading } from "MB/Components/Loading";
import { getAsyncLoadedComponent } from "MB/Components/AsyncLoadedComponent";

import { HostNavigationService } from "VSS/SDK/Services/Navigation";

import { UrlActions } from "./Constants";
import { BugBashView } from "./Components/BugBashView";
import * as AllBugBashesView_Async from "./Components/AllBugBashesView";

export enum AppViewMode {
    All,
    New,
    Results,
    Edit,
    Charts
}

export interface IAppState extends IBaseComponentState {
    appViewMode: AppViewMode;
    bugBashId?: string;
}

const AsyncAllBugBashView = getAsyncLoadedComponent(
    ["scripts/AllBugBashesView"],
    (m: typeof AllBugBashesView_Async) => m.AllBugBashesView,
    () => <Loading />);

export class App extends BaseComponent<IBaseComponentProps, IAppState> {
    private _navigationService: HostNavigationService;

    protected initializeState() {
        this.state = {
            appViewMode: null
        };
    }

    public componentDidMount() { 
        super.componentDidMount();
        this._initialize();
    }

    public componentWillUnmount() { 
        super.componentWillUnmount();
        this._detachNavigate();
    }

    public render(): JSX.Element {
        let view;

        if (this.state.appViewMode == null) {
            view = <Loading />;
        }
        else {
            switch (this.state.appViewMode) {            
                case AppViewMode.All:
                    view = <AsyncAllBugBashView />;
                    break;
                case AppViewMode.New:
                case AppViewMode.Edit:
                    view = <BugBashView pivotKey={UrlActions.ACTION_EDIT} bugBashId={this.state.bugBashId} />;
                    break;
                case AppViewMode.Results:
                    view = <BugBashView pivotKey={UrlActions.ACTION_RESULTS} bugBashId={this.state.bugBashId} />;
                    break;
                case AppViewMode.Charts:
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
        this._navigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
        this._attachNavigate();    
        
        const state = await this._navigationService.getCurrentState();
        if (!state.action) {
            this._navigationService.updateHistoryEntry(UrlActions.ACTION_ALL, null, true);
        }        
    }

    private _attachNavigate() {
        this._navigationService.attachNavigate(UrlActions.ACTION_ALL, this._navigateToHome, true);
        this._navigationService.attachNavigate(UrlActions.ACTION_EDIT, this._navigateToEditor, true);
        this._navigationService.attachNavigate(UrlActions.ACTION_RESULTS, this._navigateToResults, true);
        this._navigationService.attachNavigate(UrlActions.ACTION_CHARTS, this._navigateToCharts, true);
    }

    private _detachNavigate() {
        if (this._navigationService) {
            this._navigationService.detachNavigate(UrlActions.ACTION_ALL, this._navigateToHome);
            this._navigationService.detachNavigate(UrlActions.ACTION_EDIT, this._navigateToEditor);
            this._navigationService.detachNavigate(UrlActions.ACTION_RESULTS, this._navigateToResults);
            this._navigationService.detachNavigate(UrlActions.ACTION_CHARTS, this._navigateToCharts);
        }        
    }

    @autobind
    private _navigateToHome() {
        this.updateState({ appViewMode: AppViewMode.All, bugBashId: null });
    }

    @autobind
    private async _navigateToEditor() {
        const state = await this._navigationService.getCurrentState();
        this.updateState({ appViewMode: AppViewMode.Edit, bugBashId: state.id || null });
    }

    @autobind
    private async _navigateToResults() {
        const state = await this._navigationService.getCurrentState();
        this.updateState({ appViewMode: AppViewMode.Results, bugBashId: state.id || null });
    }

    @autobind
    private async _navigateToCharts() {
        const state = await this._navigationService.getCurrentState();
        this.updateState({ appViewMode: AppViewMode.Charts, bugBashId: state.id || null });
    }
}

export function init() {
    ReactDOM.render(<App />, $("#ext-container")[0]);
}