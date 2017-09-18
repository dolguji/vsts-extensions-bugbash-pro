import "../../css/BugBashView.scss";

import * as React from "react";

import { CoreUtils } from "MB/Utils/Core";
import { getAsyncLoadedComponent } from "MB/Components/AsyncLoadedComponent";
import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "MB/Components/BaseComponent";
import { Loading } from "MB/Components/Loading";
import { BaseStore } from "MB/Flux/Stores/BaseStore";
import { Hub, FilterPosition } from "MB/Components/Hub";

import { autobind } from "OfficeFabric/Utilities";
import { Link } from "OfficeFabric/Link";
import { Overlay } from "OfficeFabric/Overlay";
import { MessageBar, MessageBarType } from "OfficeFabric/MessageBar";
import { IContextualMenuItem } from "OfficeFabric/components/ContextualMenu/ContextualMenu.Props";

import { StoresHub } from "../Stores/StoresHub";
import { confirmAction, getBugBashUrl, navigate } from "../Helpers";
import { BugBashActions } from "../Actions/BugBashActions";
import { BugBashItemActions } from "../Actions/BugBashItemActions";
import { UrlActions, ChartsView, ResultsView, BugBashFieldNames, BugBashItemFieldNames, ErrorKeys } from "../Constants";
import { BugBash } from "../ViewModels/BugBash";
import { LongText } from "../ViewModels/LongText";
import { BugBashClientActionsHub } from "../Actions/ActionsHub";
import { BugBashErrorMessageActions } from "../Actions/BugBashErrorMessageActions";
import { BugBashDetails } from "./BugBashDetails";
import * as BugBashEditor_Async from "./BugBashEditor";
import * as BugBashResults_Async from "./BugBashResults";
import * as BugBashCharts_Async from "./BugBashCharts";

export interface IBugBashViewProps extends IBaseComponentProps {
    bugBashId?: string;
    pivotKey: string;
}

export interface IBugBashViewState extends IBaseComponentState {
    bugBash: BugBash;
    bugBashDetails?: LongText;
    selectedPivot?: string;
    filterText?: string;
    selectedChartsView?: string;
    isDetailsInEditMode?: boolean;
    selectedResultsView?: string;
    pendingItemsCount?: number;
    acceptedItemsCount?: number;
    rejectedItemsCount?: number;
}

const AsyncBugBashEditor = getAsyncLoadedComponent(
    ["scripts/BugBashEditor"],
    (m: typeof BugBashEditor_Async) => m.BugBashEditor,
    () => <Loading />);

const AsyncBugBashResults = getAsyncLoadedComponent(
    ["scripts/BugBashResults"],
    (m: typeof BugBashResults_Async) => m.BugBashResults,
    () => <Loading />);

const AsyncBugBashCharts = getAsyncLoadedComponent(
    ["scripts/BugBashCharts"],
    (m: typeof BugBashCharts_Async) => m.BugBashCharts,
    () => <Loading />);

export class BugBashView extends BaseComponent<IBugBashViewProps, IBugBashViewState> {
    private _filterChangeDelayedFunction: CoreUtils.DelayedFunction;

    protected initializeState() {
        this.state = {
            bugBash: this.props.bugBashId ? null : StoresHub.bugBashStore.getNewBugBash(),
            loading: this.props.bugBashId ? true : false,
            selectedPivot: this.props.pivotKey,
            pendingItemsCount: 0,
            acceptedItemsCount: 0,
            rejectedItemsCount: 0,
            isDetailsInEditMode: false
        };
    }

    protected getStores(): BaseStore<any, any, any>[] {
        return [StoresHub.bugBashStore, StoresHub.bugBashItemStore, StoresHub.longTextStore];
    }

    protected getStoresState(): IBugBashViewState {
        const bugBashItems = StoresHub.bugBashItemStore.getBugBashItems(this.props.bugBashId || "");

        return {
            loading: this.props.bugBashId ? StoresHub.bugBashStore.isLoading(this.props.bugBashId) : StoresHub.bugBashStore.isLoading(),
            bugBash: this.props.bugBashId ? StoresHub.bugBashStore.getItem(this.props.bugBashId) : StoresHub.bugBashStore.getNewBugBash(),
            bugBashDetails: this.props.bugBashId ? StoresHub.longTextStore.getItem(this.props.bugBashId) : null,
            pendingItemsCount: bugBashItems ? bugBashItems.filter(b => !b.isAccepted && !b.getFieldValue<boolean>(BugBashItemFieldNames.Rejected, true)).length : 0,
            acceptedItemsCount: bugBashItems ? bugBashItems.filter(b => b.isAccepted).length : 0,
            rejectedItemsCount: bugBashItems ? bugBashItems.filter(b => !b.isAccepted && b.getFieldValue<boolean>(BugBashItemFieldNames.Rejected, true)).length : 0
        } as IBugBashViewState;
    }

    public componentDidMount() {
        super.componentDidMount();
        
        if (this.props.bugBashId) {
            BugBashActions.initializeBugBash(this.props.bugBashId);
        }
    }  

    public componentWillUnmount() {
        super.componentWillUnmount();

        if (this.state.bugBash) {
            this.state.bugBash.reset(false);
        }

        const items = StoresHub.bugBashItemStore.getBugBashItems(this.props.bugBashId || "") || [];
        for (const item of items) {
            item.reset(false);
        }

        StoresHub.bugBashItemStore.getNewBugBashItem().reset(false);

        if (this.state.bugBashDetails) {
            this.state.bugBashDetails.reset();
        }

        BugBashErrorMessageActions.dismissErrorMessage(ErrorKeys.BugBashError);
        BugBashErrorMessageActions.dismissErrorMessage(ErrorKeys.BugBashItemError);
        BugBashErrorMessageActions.dismissErrorMessage(ErrorKeys.BugBashDetailsError);
    }
    
    public componentWillReceiveProps(nextProps: Readonly<IBugBashViewProps>) {
        if (nextProps.bugBashId !== this.props.bugBashId) {
            BugBashErrorMessageActions.dismissErrorMessage(ErrorKeys.BugBashError);
            BugBashErrorMessageActions.dismissErrorMessage(ErrorKeys.BugBashItemError);
            BugBashErrorMessageActions.dismissErrorMessage(ErrorKeys.BugBashDetailsError);

            if (!nextProps.bugBashId) {
                this.updateState({
                    bugBash: StoresHub.bugBashStore.getNewBugBash(),
                    loading: false,
                    selectedPivot: nextProps.pivotKey
                });
            }
            else if (StoresHub.bugBashStore.getItem(nextProps.bugBashId)) {
                this.updateState({
                    bugBash: StoresHub.bugBashStore.getItem(nextProps.bugBashId),
                    loading: false,
                    selectedPivot: nextProps.pivotKey
                });
            } 
            else {
                this.updateState({
                    bugBash: null,
                    loading: true,
                    selectedPivot: nextProps.pivotKey
                });

                BugBashActions.initializeBugBash(nextProps.bugBashId);
            }
        }
        else {
            this.updateState({
                selectedPivot: nextProps.pivotKey
            } as IBugBashViewState);
        }
    }

    public render(): JSX.Element {
        return <div className="bugbash-view">
            { this.state.loading && <Overlay className="loading-overlay"><Loading /></Overlay> }
            { this.state.bugBash && this._renderHub() }
        </div>;
    }

    @autobind
    private _onTitleRender(): React.ReactNode {
        const bugBash = this.state.bugBash;
        
        let title = bugBash.getFieldValue<string>(BugBashFieldNames.Title);
        let className = "bugbash-title";
        if (bugBash.isDirty()) {
            className += " is-dirty";
            title = "* " + title;
        }

        if (!bugBash.isNew() && !bugBash.isValid()) {
            className += " is-invalid";
        }
        
        return (
            <div className="bugbash-hub-title">
                <Link 
                    href={getBugBashUrl(null, UrlActions.ACTION_ALL)}
                    onClick={async (e: React.MouseEvent<HTMLElement>) => {
                        if (!e.ctrlKey) {
                            e.preventDefault();
                            
                            const confirm = await confirmAction(this._isAnyProviderDirty(), 
                                "You have unsaved changes in the bug bash. Navigating to Home will revert your changes. Are you sure you want to do that?");

                            if (confirm) {
                                navigate(UrlActions.ACTION_ALL, null);
                            } 
                        }
                    }}
                    className="bugbashes-button">
                    Bug Bashes
                </Link>
                <span className="divider">></span>
                <span className={className}>{title}</span>
            </div>
        )
    }

    private _isAnyProviderDirty(): boolean {
        const bugBash = this.state.bugBash;
        const items = StoresHub.bugBashItemStore.getBugBashItems(this.props.bugBashId || "") || [];
        const isAnyBugBashItemDirty = items.some(item => item.isDirty());
        return bugBash.isDirty() || isAnyBugBashItemDirty || StoresHub.bugBashItemStore.getNewBugBashItem().isDirty() || (this.state.bugBashDetails && this.state.bugBashDetails.isDirty());
    }

    private _renderHub(): React.ReactNode {
        return <Hub 
            className="bugbash-hub"
            onTitleRender={this._onTitleRender}
            pivotProps={{
                onPivotClick: (selectedPivotKey: string) => {
                    this.updateState({selectedPivot: selectedPivotKey} as IBugBashViewState);
                    navigate(selectedPivotKey, null, false, true, null, true);                  
                },
                initialSelectedKey: this.state.selectedPivot,
                onRenderPivotContent: (key: string) => {
                    const extraCss = this.state.bugBash.isNew() ? "new-bugbash" : "";

                    switch (key) {
                        case "edit":
                            return <div className="bugbash-hub-contents bugbash-editor-hub-contents">
                                {this._renderEditor()}
                            </div>;
                        case "results":
                            return <div className={`bugbash-hub-contents bugbash-results-hub-contents ${extraCss}`}>
                                {this._renderResults()}
                            </div>;
                        case "charts":
                            return <div className="bugbash-hub-contents bugbash-charts-hub-contents">
                                {this._renderCharts()}
                            </div>;
                        case "details":
                            return <div className={`bugbash-hub-contents bugbash-details-hub-contents ${extraCss}`}>
                                {this._renderDetails()}
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
                    },
                    {
                        key: "details",
                        text: "Details",
                        commands: this._getDetailsViewCommands()
                    }
                ]
            }}
        />;
    }    

    private _renderEditor(): JSX.Element {
        return <AsyncBugBashEditor
            bugBash={this.state.bugBash} />;
    }
    
    private _renderResults(): JSX.Element {
        if (!this.state.bugBash.isNew()) {
            return <AsyncBugBashResults
                filterText={this.state.filterText || ""}
                view={this.state.selectedResultsView}
                bugBash={this.state.bugBash} />;
        }
        else {
            return <MessageBar messageBarType={MessageBarType.info} className="message-panel">
                Please save the bug bash first to view results.
            </MessageBar>;
        }
    }

    private _renderCharts(): JSX.Element {
        if (!this.state.bugBash.isNew()) {
            return <AsyncBugBashCharts
                bugBash={this.state.bugBash}
                view={this.state.selectedChartsView} />;
        }
        else {
            return <MessageBar messageBarType={MessageBarType.info} className="message-panel">
                Please save the bug bash first to view charts.
            </MessageBar>;
        }
    }

    private _renderDetails(): JSX.Element {
        if (!this.state.bugBash.isNew()) {
            return <BugBashDetails
                isEditMode={this.state.isDetailsInEditMode}
                id={this.props.bugBashId} />;
        }
        else {
            return <MessageBar messageBarType={MessageBarType.info} className="message-panel">
                Please save the bug bash first to view and edit its details.
            </MessageBar>;
        }
    }

    private _getEditorViewCommands(): IContextualMenuItem[] {
        return [
            {
                key: "save", name: "Save", iconProps: {iconName: "Save"}, 
                disabled: !this.state.bugBash.isDirty() || !this.state.bugBash.isValid(),
                onClick: this._saveBugBash
            },
            {
                key: "undo", name: "Undo", iconProps: {iconName: "Undo"}, 
                disabled: this.state.bugBash.isNew() || !this.state.bugBash.isDirty(),
                onClick: this._revertBugBash
            },
            {
                key: "refresh", name: "Refresh", iconProps: {iconName: "Refresh"}, 
                disabled: this.state.bugBash.isNew(),
                onClick: this._refreshBugBash
            }
        ];
    }

    private _getDetailsViewCommands(): IContextualMenuItem[] {
        if (this.state.isDetailsInEditMode) {
            return [
                {
                    key: "save", name: "Save", iconProps: {iconName: "Save"}, 
                    disabled: !this.state.bugBashDetails || !this.state.bugBashDetails.isDirty(),
                    onClick: this._saveBugBashDetails
                },
                {
                    key: "undo", name: "Undo", iconProps: {iconName: "Undo"}, 
                    disabled: !this.state.bugBashDetails || !this.state.bugBashDetails.isDirty(),
                    onClick: this._revertBugBashDetails
                },
                {
                    key: "refresh", name: "Refresh", iconProps: {iconName: "Refresh"}, 
                    disabled: !this.state.bugBashDetails,
                    onClick: this._refreshBugBashDetails
                },
                {
                    key: "done", name: "Done", iconProps: {iconName: "Accept"}, 
                    disabled: !this.state.bugBashDetails,
                    onClick: () => this.updateState({isDetailsInEditMode: false} as IBugBashViewState)
                },
            ];
        }
        else {
            return [
                {
                    key: "edit", name: "Edit", iconProps: {iconName: "Edit"}, 
                    disabled: !this.state.bugBashDetails,
                    onClick: () => this.updateState({isDetailsInEditMode: true} as IBugBashViewState)
                },
                {
                    key: "refresh", name: "Refresh", iconProps: {iconName: "Refresh"}, 
                    disabled: !this.state.bugBashDetails,
                    onClick: this._refreshBugBashDetails
                }
            ];
        }
    }

    private _getResultViewCommands(): IContextualMenuItem[] {
        return [
            {
                key: "refresh", name: "Refresh", iconProps: {iconName: "Refresh"}, 
                disabled: this.state.bugBash.isNew(),
                onClick: this._refreshBugBashItems
            },
            {
                key: "newitem", name: "New Item", iconProps: {iconName: "Add"}, 
                disabled: this.state.bugBash.isNew(),
                onClick: () => {
                    StoresHub.bugBashItemStore.getNewBugBashItem().reset(false);
                    BugBashClientActionsHub.SelectedBugBashItemChanged.invoke(null);
                }
            }            
        ];
    }

    private _getResultCount(): number {
        switch (this.state.selectedResultsView) {
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
                key: "resultsview", name: this.state.selectedResultsView || ResultsView.PendingItemsOnly, 
                iconProps: { iconName: "Equalizer" },
                disabled: this.state.bugBash.isNew(),
                subMenuProps: {
                    items: [                        
                        {
                            key: "pending", name: ResultsView.PendingItemsOnly, canCheck: true,
                            checked: !this.state.selectedResultsView || this.state.selectedResultsView === ResultsView.PendingItemsOnly
                        },
                        {
                            key: "rejected", name: ResultsView.RejectedItemsOnly, canCheck: true,
                            checked: this.state.selectedResultsView === ResultsView.RejectedItemsOnly
                        },
                        {
                            key: "accepted", name: ResultsView.AcceptedItemsOnly, canCheck: true,
                            checked: this.state.selectedResultsView === ResultsView.AcceptedItemsOnly
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
                disabled: this.state.bugBash.isNew(),
                onClick: this._refreshBugBashItems
            }            
        ];
    }

    private _getChartsViewFarCommands(): IContextualMenuItem[] {
        return [
            {
                key: "chartsview", name: this.state.selectedChartsView || ChartsView.All, 
                iconProps: { iconName: "Equalizer" },
                disabled: this.state.bugBash.isNew(),
                subMenuProps: {
                    items: [
                        {
                            key: "all", name: ChartsView.All, canCheck: true,
                            checked: !this.state.selectedChartsView || this.state.selectedChartsView === ChartsView.All
                        },
                        {
                            key: "pending", name: ChartsView.PendingItemsOnly, canCheck: true,
                            checked: this.state.selectedChartsView === ChartsView.PendingItemsOnly
                        },
                        {
                            key: "rejected", name: ChartsView.RejectedItemsOnly, canCheck: true,
                            checked: this.state.selectedChartsView === ChartsView.RejectedItemsOnly
                        },
                        {
                            key: "accepted", name: ChartsView.AcceptedItemsOnly, canCheck: true,
                            checked: this.state.selectedChartsView === ChartsView.AcceptedItemsOnly
                        },
                    ],
                    onItemClick: this._onChangeChartsView
                }
            }
        ];
    }

    @autobind
    private _onChangeChartsView(_ev?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem) {
        this.updateState({selectedChartsView: item.name} as IBugBashViewState);
    }

    @autobind
    private _onChangeResultsView(_ev?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem) {
        this.updateState({selectedResultsView: item.name} as IBugBashViewState);
    }

    @autobind
    private _onFilterTextChange(filterText: string) {
        if (this._filterChangeDelayedFunction) {
            this._filterChangeDelayedFunction.cancel();
        }

        this._filterChangeDelayedFunction = CoreUtils.delay(this, 200, () => {
            this.updateState({filterText: filterText} as IBugBashViewState);
        });        
    }

    @autobind
    private async _refreshBugBashItems() {
        const items = StoresHub.bugBashItemStore.getBugBashItems(this.props.bugBashId) || [];
        const isAnyBugBashItemDirty = items.some(item => item.isDirty());
        const confirm = await confirmAction(isAnyBugBashItemDirty || StoresHub.bugBashItemStore.getNewBugBashItem().isDirty(), "You have some unsaved items in the list. Refreshing the page will remove all the unsaved data. Are you sure you want to do it?");
        if (confirm) {
            StoresHub.bugBashItemStore.getNewBugBashItem().reset(false);            
            await BugBashItemActions.refreshItems(this.props.bugBashId);
            StoresHub.workItemStore.clearStore();
            BugBashClientActionsHub.SelectedBugBashItemChanged.invoke(null);
        }
    }

    @autobind
    private async _revertBugBash() {
        const confirm = await confirmAction(true, "Are you sure you want to undo your changes to this instance?");
        if (confirm) {
            this.state.bugBash.reset();
        }
    }

    @autobind
    private async _refreshBugBash() {
        const confirm = await confirmAction(this.state.bugBash.isDirty(), 
            "Refreshing will undo your unsaved changes. Are you sure you want to do that?");

        if (confirm) {
            this.state.bugBash.refresh();
        }
    }

    @autobind
    private _saveBugBash() {
        this.state.bugBash.save();
    }

    @autobind
    private async _revertBugBashDetails() {
        const confirm = await confirmAction(true, "Are you sure you want to undo your changes to this instance?");
        if (confirm) {
            this.state.bugBashDetails.reset();
        }
    }

    @autobind
    private async _refreshBugBashDetails() {
        const confirm = await confirmAction(this.state.bugBash.isDirty(), 
            "Refreshing will undo your unsaved changes. Are you sure you want to do that?");

        if (confirm) {
            this.state.bugBashDetails.refresh();
        }
    }

    @autobind
    private _saveBugBashDetails() {
        this.state.bugBashDetails.save();
    }
}