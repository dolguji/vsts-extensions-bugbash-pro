import "../css/App.scss";

import * as React from "react";
import * as ReactDOM from "react-dom";

import { HostNavigationService } from "VSS/SDK/Services/Navigation";

import { Fabric } from "OfficeFabric/Fabric";
import { autobind } from "OfficeFabric/Utilities";
import { initializeIcons } from "@uifabric/icons";

import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "MB/Components/BaseComponent";
import { Loading } from "MB/Components/Loading";
import { getAsyncLoadedComponent } from "MB/Components/AsyncLoadedComponent";
import { Badge } from "MB/Components/Badge";
import { BaseStore } from "MB/Flux/Stores/BaseStore";

import { UrlActions } from "./Constants";
import { BugBashView } from "./Components/BugBashView";
import * as AllBugBashesView_Async from "./Components/AllBugBashesView";
import { navigate } from "./Helpers";
import { StoresHub } from "./Stores/StoresHub";
import { SettingsActions } from "./Actions/SettingsActions";

export enum AppViewMode {
    All,
    New,
    Results,
    Edit,
    Charts,
    Details
}

export interface IAppState extends IBaseComponentState {
    appViewMode: AppViewMode;
    bugBashId?: string;
    userSettingsAvailable?: boolean;
}

const AsyncAllBugBashView = getAsyncLoadedComponent(
    ["scripts/AllBugBashesView"],
    (m: typeof AllBugBashesView_Async) => m.AllBugBashesView,
    () => <Loading />);

export class App extends BaseComponent<IBaseComponentProps, IAppState> {
    private _navigationService: HostNavigationService;

    protected initializeState() {
        this.state = {
            appViewMode: null,
            userSettingsAvailable: true
        };
    }

    protected getStores(): BaseStore<any, any, any>[] {
        return [StoresHub.userSettingsStore];
    }

    public componentDidMount() { 
        super.componentDidMount();
        this._initialize();
        SettingsActions.initializeUserSettings();
    }

    public componentWillUnmount() { 
        super.componentWillUnmount();
        this._detachNavigate();
    }

    protected getStoresState(): IAppState {
        return {            
            userSettingsAvailable: StoresHub.userSettingsStore.isLoading() ? true : StoresHub.userSettingsStore.getItem(VSS.getWebContext().user.email) != null
        } as IAppState;
    }

    public shouldComponentUpdate(_nextProps: Readonly<IBaseComponentProps>, nextState: Readonly<IAppState>) {
        if (this.state.appViewMode !== nextState.appViewMode
            || this.state.bugBashId !== nextState.bugBashId
            || this.state.userSettingsAvailable !== nextState.userSettingsAvailable
            || this.state.loading !== nextState.loading) {

            return true;
        }

        return false;
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
                case AppViewMode.Details:
                    view = <BugBashView pivotKey={UrlActions.ACTION_DETAILS} bugBashId={this.state.bugBashId} />;
                    break;
                default:
                    view = <Loading />;
                    break;
            }
        }

        return (
            <Fabric className="fabric-container">
                { this._renderBadge() }
                { view }
            </Fabric>
        );
    }
    
    private _renderBadge(): JSX.Element {
        if (!this.state.userSettingsAvailable) {
            return <Badge className="bugbash-badge" notificationCount={1}>
                <div className="bugbash-badge-callout">
                    <div className="badge-callout-header">
                        Don't forget to set your associated team!!
                    </div>
                    <div className="badge-callout-inner">
                        <div>
                            You can set a team associated with you in the current project by clicking on "Settings" link in the Bug Bash home page.
                        </div>
                        <div>
                            If you have set a team associated with you, any bug bash item created by you will also count towards your team.
                        </div>
                        <div>
                            This will be reflected in the "Created By" chart in a Bug Bash.
                        </div>
                    </div>
                </div>
            </Badge>;
        }

        return null;
    }

    private async _initialize() {
        this._navigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
        this._attachNavigate();    
        
        const state = await this._navigationService.getCurrentState();
        if (!state.action) {
            navigate(UrlActions.ACTION_ALL, null, true);
        }        
    }

    private _attachNavigate() {
        this._navigationService.attachNavigate(UrlActions.ACTION_ALL, this._navigateToHome, true);
        this._navigationService.attachNavigate(UrlActions.ACTION_EDIT, this._navigateToEditor, true);
        this._navigationService.attachNavigate(UrlActions.ACTION_RESULTS, this._navigateToResults, true);
        this._navigationService.attachNavigate(UrlActions.ACTION_CHARTS, this._navigateToCharts, true);
        this._navigationService.attachNavigate(UrlActions.ACTION_DETAILS, this._navigateToDetails, true);
    }

    private _detachNavigate() {
        if (this._navigationService) {
            this._navigationService.detachNavigate(UrlActions.ACTION_ALL, this._navigateToHome);
            this._navigationService.detachNavigate(UrlActions.ACTION_EDIT, this._navigateToEditor);
            this._navigationService.detachNavigate(UrlActions.ACTION_RESULTS, this._navigateToResults);
            this._navigationService.detachNavigate(UrlActions.ACTION_CHARTS, this._navigateToCharts);
            this._navigationService.detachNavigate(UrlActions.ACTION_DETAILS, this._navigateToDetails);
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

    @autobind
    private async _navigateToDetails() {
        const state = await this._navigationService.getCurrentState();
        this.updateState({ appViewMode: AppViewMode.Details, bugBashId: state.id || null });
    }
}

export function init() {
    initializeIcons();    
    ReactDOM.render(<App />, $("#ext-container")[0]);
}