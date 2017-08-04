import "../../css/BugBashView.scss";

import * as React from "react";
import { HostNavigationService } from "VSS/SDK/Services/Navigation";
import * as EventsService from "VSS/Events/Services";
import Context = require("VSS/Context");

import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { LazyLoad } from "VSTS_Extension/Components/Common/LazyLoad";
import { Loading } from "VSTS_Extension/Components/Common/Loading";
import { BaseStore } from "VSTS_Extension/Flux/Stores/BaseStore";
import { Hub, FilterPosition } from "VSTS_Extension/Components/Common/Hub/Hub";

import { autobind } from "OfficeFabric/Utilities";
import { Link } from "OfficeFabric/Link";
import { Overlay } from "OfficeFabric/Overlay";
import { MessageBar, MessageBarType } from "OfficeFabric/MessageBar";
import { IContextualMenuItem } from "OfficeFabric/components/ContextualMenu/ContextualMenu.Props";

import { IBugBashViewModel, IBugBash } from "../Interfaces";
import { StoresHub } from "../Stores/StoresHub";
import { confirmAction, BugBashHelpers, BugBashItemHelpers } from "../Helpers";
import { BugBashActions } from "../Actions/BugBashActions";
import { BugBashItemActions } from "../Actions/BugBashItemActions";
import { BugBashItemCommentActions } from "../Actions/BugBashItemCommentActions";
import { UrlActions, Events, ChartsView, ResultsView } from "../Constants";
import { BugBashErrorMessageActions } from "../Actions/BugBashErrorMessageActions";

export interface IBugBashViewProps extends IBaseComponentProps {
    bugBashId?: string;
    pivotKey: string;
}

export interface IBugBashViewState extends IBaseComponentState {
    bugBashViewModel: IBugBashViewModel;
    selectedPivot?: string;
    isAnyBugBashItemDirty?: boolean;
    bugBashEditorError?: string;
    filterText?: string;
    chartsView?: string;
    resultsView?: string;
    pendingItemsCount?: number;
    acceptedItemsCount?: number;
    rejectedItemsCount?: number;
}

export class BugBashView extends BaseComponent<IBugBashViewProps, IBugBashViewState> {
    protected initializeState() {
        this.state = {
            bugBashViewModel: this.props.bugBashId ? null : BugBashHelpers.getNewViewModel(),
            loading: this.props.bugBashId ? true : false,
            selectedPivot: this.props.pivotKey,
            pendingItemsCount: 0,
            acceptedItemsCount: 0,
            rejectedItemsCount: 0,
        };
    }

    protected getStores(): BaseStore<any, any, any>[] {
        return [StoresHub.bugBashStore, StoresHub.bugBashItemStore];
    }

    protected getStoresState(): IBugBashViewState {
        const bugBashItems = StoresHub.bugBashItemStore.getBugBashItems(this.props.bugBashId);

        return {
            loading: StoresHub.bugBashStore.isLoading(this.props.bugBashId),
            bugBashEditorError: null,
            bugBashViewModel: BugBashHelpers.getViewModel(StoresHub.bugBashStore.getItem(this.props.bugBashId)),
            pendingItemsCount: bugBashItems ? bugBashItems.filter(b => !BugBashItemHelpers.isAccepted(b) && !b.rejected).length : 0,
            acceptedItemsCount: bugBashItems ? bugBashItems.filter(b => BugBashItemHelpers.isAccepted(b)).length : 0,
            rejectedItemsCount: bugBashItems ? bugBashItems.filter(b => !BugBashItemHelpers.isAccepted(b) && b.rejected).length : 0
        } as IBugBashViewState;
    }

    public componentDidMount(): void {
        super.componentDidMount();
        
        if (this.props.bugBashId) {
            this._initializeBugBash(this.props.bugBashId);
        }
    }  
    
    public componentWillReceiveProps(nextProps: Readonly<IBugBashViewProps>): void {
        if (nextProps.bugBashId !== this.props.bugBashId) {
            if (!nextProps.bugBashId) {
                this.updateState({
                    bugBashViewModel: BugBashHelpers.getNewViewModel(),
                    loading: false,
                    selectedPivot: nextProps.pivotKey
                });                
            }
            else if (StoresHub.bugBashStore.getItem(nextProps.bugBashId)) {
                this.updateState({
                    bugBashViewModel: BugBashHelpers.getViewModel(StoresHub.bugBashStore.getItem(nextProps.bugBashId)),
                    loading: false,
                    selectedPivot: nextProps.pivotKey
                });
            } 
            else {
                this.updateState({
                    bugBashViewModel: null,
                    loading: true,
                    selectedPivot: nextProps.pivotKey
                });

                this._initializeBugBash(nextProps.bugBashId);
            }
        }
        else {
            this.updateState({
                selectedPivot: nextProps.pivotKey
            } as IBugBashViewState);
        }
    }

    private async _initializeBugBash(bugBashId: string) {        
        try {
            await BugBashActions.initializeBugBash(bugBashId);
        }
        catch (e) {
            BugBashErrorMessageActions.showErrorMessage(e);
            let navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
            navigationService.updateHistoryEntry(UrlActions.ACTION_ALL, null, true);
        }
    }

    public render(): JSX.Element {
        return <div className="bugbash-view">
            { this.state.loading && <Overlay className="loading-overlay"><Loading /></Overlay> }
            { this.state.bugBashViewModel && this._renderHub() }
        </div>;
    }

    private _onTitleRender(title: string): React.ReactNode {
        const pageContext = Context.getPageContext();
        const navigation = pageContext.navigation;
        const webContext = VSS.getWebContext();
        const bugBashUrl = `${webContext.collection.uri}/${webContext.project.name}/_${navigation.currentController}/${navigation.currentAction}/${navigation.currentParameters}#_a=${UrlActions.ACTION_ALL}`;
        return (
            <div className="bugbash-hub-title">
                <Link 
                    href={bugBashUrl}
                    onClick={async (e: React.MouseEvent<HTMLElement>) => {
                        if (!e.ctrlKey) {
                            e.preventDefault();
                            const confirm = await confirmAction(BugBashHelpers.isDirty(this.state.bugBashViewModel), "You have unsaved changes in the bug bash. Navigating to Home will revert your changes. Are you sure you want to do that?");

                            if (confirm) {
                                let navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
                                navigationService.updateHistoryEntry(UrlActions.ACTION_ALL, null);
                            } 
                        }
                    }}
                    className="bugbashes-button">
                    Bug Bashes
                </Link>
                <span className="divider">></span>
                <span className="bugbash-title">{title}</span>
            </div>
        )
    }

    private _renderHub(): React.ReactNode {
        const updatedBugBash = this.state.bugBashViewModel.updatedBugBash;

        if (updatedBugBash) {
            let title = updatedBugBash.title;
            let className = "bugbash-hub";
            if (BugBashHelpers.isDirty(this.state.bugBashViewModel)) {
                className += " is-dirty";
                title = "* " + title;
            }

            return <Hub 
                className={className}
                onTitleRender={() => this._onTitleRender(title)}
                pivotProps={{
                    onPivotClick: async (selectedPivotKey: string) => {
                        this.updateState({selectedPivot: selectedPivotKey} as IBugBashViewState);
                        let navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
                        navigationService.updateHistoryEntry(selectedPivotKey, null, false, true, null, true);
                    },
                    initialSelectedKey: this.state.selectedPivot,
                    onRenderPivotContent: (key: string) => {
                        switch (key) {
                            case "edit":
                                return <div className="bugbash-hub-contents bugbash-editor-hub-contents">
                                    {this._renderEditor()}
                                </div>;
                            case "results":
                                return <div className="bugbash-hub-contents bugbash-results-hub-contents">
                                    {this._renderResults()}
                                </div>;
                            case "charts":
                                return <div className="bugbash-hub-contents bugbash-charts-hub-contents">
                                    {this._renderCharts()}
                                </div>;
                        }
                    },
                    pivots: [
                        {
                            key: "results",
                            text: "Results",
                            commands: this._getResultViewCommands(),
                            farCommands: this._getResultViewFarCommands(),
                            filterProps: {
                                showFilter: true,
                                onFilterChange: this._onFilterTextChange,
                                filterPosition: FilterPosition.Left
                            }
                        },
                        {
                            key: "edit",
                            text: "Editor",
                            commands: this._getEditorViewCommands(),
                        },
                        {
                            key: "charts",
                            text: "Charts",
                            commands: this._getChartsViewCommands(),
                            farCommands: this._getChartsViewFarCommands()
                        }
                    ]
                }}
            />;
        }
        else {
            return null;
        }
    }    

    private _renderEditor(): JSX.Element {
        return <LazyLoad module="scripts/BugBashEditor">
            {(BugBashEditor) => (
                <BugBashEditor.BugBashEditor 
                    bugBash={{...this.state.bugBashViewModel.updatedBugBash}}
                    error={this.state.bugBashEditorError}
                    onChange={this._onBugBashChange}
                    save={this._saveBugBash}
                />
            )}
        </LazyLoad>;
    }
    
    private _renderResults(): JSX.Element {
        if (!BugBashHelpers.isNew(this.state.bugBashViewModel.originalBugBash)) {
            return <LazyLoad module="scripts/BugBashResults">
                {(BugBashResults) => (
                    <BugBashResults.BugBashResults 
                        filterText={this.state.filterText || ""}
                        view={this.state.resultsView}
                        bugBash={{...this.state.bugBashViewModel.originalBugBash}} />
                )}
            </LazyLoad>;
        }
        else {
            return <MessageBar messageBarType={MessageBarType.info} className="message-panel">
                Please save the bug bash first to view results.
            </MessageBar>;
        }
    }

    private _renderCharts(): JSX.Element {
        if (!BugBashHelpers.isNew(this.state.bugBashViewModel.originalBugBash)) {
            return <LazyLoad module="scripts/BugBashCharts">
                {(BugBashCharts) => (
                    <BugBashCharts.BugBashCharts 
                        bugBash={{...this.state.bugBashViewModel.originalBugBash}} 
                        view={this.state.chartsView} />
                )}
            </LazyLoad>;
        }
        else {
            return <MessageBar messageBarType={MessageBarType.info} className="message-panel">
                Please save the bug bash first to view charts.
            </MessageBar>;
        }
    }

    private _getEditorViewCommands(): IContextualMenuItem[] {
        return [
            {
                key: "save", name: "Save", iconProps: {iconName: "Save"}, 
                disabled: !BugBashHelpers.isDirty(this.state.bugBashViewModel) || !BugBashHelpers.isValid(this.state.bugBashViewModel),
                onClick: this._saveBugBash
            },
            {
                key: "undo", name: "Undo", iconProps: {iconName: "Undo"}, 
                disabled: BugBashHelpers.isNew(this.state.bugBashViewModel.originalBugBash) || !BugBashHelpers.isDirty(this.state.bugBashViewModel),
                onClick: this._revertBugBash
            },
            {
                key: "refresh", name: "Refresh", iconProps: {iconName: "Refresh"}, 
                disabled: BugBashHelpers.isNew(this.state.bugBashViewModel.originalBugBash),
                onClick: this._refreshBugBash
            }
        ];
    }

    private _getResultViewCommands(): IContextualMenuItem[] {
        return [
            {
                key: "refresh", name: "Refresh", iconProps: {iconName: "Refresh"}, 
                disabled: BugBashHelpers.isNew(this.state.bugBashViewModel.originalBugBash),
                onClick: this._refreshBugBashItems
            },
            {
                key: "newitem", name: "New Item", iconProps: {iconName: "Add"}, 
                disabled: BugBashHelpers.isNew(this.state.bugBashViewModel.originalBugBash),
                onClick: () => {
                    EventsService.getService().fire(Events.RefreshItems);
                }
            }            
        ];
    }

    private _getResultCount(): number {
        switch (this.state.resultsView) {
            case ResultsView.RejectedItemsOnly:
                return this.state.rejectedItemsCount;
            case ResultsView.AcceptedItemsOnly:
                return this.state.acceptedItemsCount;
            default:
                return this.state.pendingItemsCount;
        }
    }

    private _getResultViewFarCommands(): IContextualMenuItem[] {
        return [
            {
                key:"resultCount", name: `${this._getResultCount()} results`
            },
            {
                key: "resultsview", name: this.state.resultsView || ResultsView.PendingItemsOnly, 
                iconProps: { iconName: "Equalizer" },
                disabled: BugBashHelpers.isNew(this.state.bugBashViewModel.originalBugBash),
                subMenuProps: {
                    items: [                        
                        {
                            key: "pending", name: ResultsView.PendingItemsOnly, canCheck: true,
                            checked: !this.state.resultsView || this.state.resultsView === ResultsView.PendingItemsOnly
                        },
                        {
                            key: "rejected", name: ResultsView.RejectedItemsOnly, canCheck: true,
                            checked: this.state.resultsView === ResultsView.RejectedItemsOnly
                        },
                        {
                            key: "accepted", name: ResultsView.AcceptedItemsOnly, canCheck: true,
                            checked: this.state.resultsView === ResultsView.AcceptedItemsOnly
                        },
                    ],
                    onItemClick: this._onChangeResultsView
                }
            }
        ];
    }

    private _getChartsViewCommands(): IContextualMenuItem[] {
        return [
            {
                key: "refresh", name: "Refresh", iconProps: {iconName: "Refresh"}, 
                disabled: BugBashHelpers.isNew(this.state.bugBashViewModel.originalBugBash),
                onClick: this._refreshBugBashItems
            }            
        ];
    }

    private _getChartsViewFarCommands(): IContextualMenuItem[] {
        return [
            {
                key: "chartsview", name: this.state.chartsView || ChartsView.All, 
                iconProps: { iconName: "Equalizer" },
                disabled: BugBashHelpers.isNew(this.state.bugBashViewModel.originalBugBash),
                subMenuProps: {
                    items: [
                        {
                            key: "all", name: ChartsView.All, canCheck: true,
                            checked: !this.state.chartsView || this.state.chartsView === ChartsView.All
                        },
                        {
                            key: "pending", name: ChartsView.PendingItemsOnly, canCheck: true,
                            checked: this.state.chartsView === ChartsView.PendingItemsOnly
                        },
                        {
                            key: "rejected", name: ChartsView.RejectedItemsOnly, canCheck: true,
                            checked: this.state.chartsView === ChartsView.RejectedItemsOnly
                        },
                        {
                            key: "accepted", name: ChartsView.AcceptedItemsOnly, canCheck: true,
                            checked: this.state.chartsView === ChartsView.AcceptedItemsOnly
                        },
                    ],
                    onItemClick: this._onChangeChartsView
                }
            }
        ];
    }

    @autobind
    private _onChangeChartsView(_ev?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem) {
        this.updateState({chartsView: item.name} as IBugBashViewState);
    }

    @autobind
    private _onChangeResultsView(_ev?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem) {
        this.updateState({resultsView: item.name} as IBugBashViewState);
    }

    @autobind
    private _onFilterTextChange(filterText: string) {
        this.updateState({filterText: filterText} as IBugBashViewState);
    }

    @autobind
    private _onBugBashChange(updatedBugBash: IBugBash) {
        let newViewModel = {...this.state.bugBashViewModel};
        newViewModel.updatedBugBash = {...updatedBugBash};
        this.updateState({
            bugBashViewModel: newViewModel
        } as IBugBashViewState);
    }

    @autobind
    private async _refreshBugBashItems() {
        const confirm = await confirmAction(this.state.isAnyBugBashItemDirty, "You have some unsaved items in the list. Refreshing the page will remove all the unsaved data. Are you sure you want to do it?");
        if (confirm) {                        
            await BugBashItemActions.refreshItems(this.props.bugBashId);
            BugBashItemCommentActions.clearComments();
            EventsService.getService().fire(Events.RefreshItems);
        }
    }

    @autobind
    private async _revertBugBash() {
        const confirm = await confirmAction(true, "Are you sure you want to undo your changes to this instance?");
        if (confirm) {
            let newViewModel = {...this.state.bugBashViewModel};
            newViewModel.updatedBugBash = {...newViewModel.originalBugBash};
            this.updateState({
                bugBashViewModel: newViewModel
            } as IBugBashViewState);
        }
    }

    @autobind
    private async _refreshBugBash() {
        const confirm = await confirmAction(BugBashHelpers.isDirty(this.state.bugBashViewModel), 
            "Refreshing the item will undo your unsaved changes. Are you sure you want to do that?");

        if (confirm) {
            try {
                await BugBashActions.refreshBugBash(this.props.bugBashId);
            }
            catch (e) {
                this.updateState({bugBashEditorError: e} as IBugBashViewState);
            }                        
        }
    }

    @autobind
    private async _saveBugBash() {        
        if (BugBashHelpers.isDirty(this.state.bugBashViewModel) && BugBashHelpers.isValid(this.state.bugBashViewModel)) {
            const originalBugBash = this.state.bugBashViewModel.originalBugBash;
            const updatedBugBash = this.state.bugBashViewModel.updatedBugBash;

            if (BugBashHelpers.isNew(originalBugBash)) {
                try {
                    const createdBugBash = await BugBashActions.createBugBash(updatedBugBash);
                    let navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
                    navigationService.updateHistoryEntry(UrlActions.ACTION_EDIT, { id: createdBugBash.id }, true);
                }
                catch (e) {
                    this.updateState({bugBashEditorError: e} as IBugBashViewState);
                }
            }
            else {
                try {
                    await BugBashActions.updateBugBash(updatedBugBash);
                }
                catch (e) {
                    this.updateState({bugBashEditorError: e} as IBugBashViewState);
                }
            }
        }        
    }
}