import "../../css/BugBashView.scss";

import * as React from "react";
import Utils_String = require("VSS/Utils/String");
import { HostNavigationService } from "VSS/SDK/Services/Navigation";

import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { LazyLoad } from "VSTS_Extension/Components/Common/LazyLoad";
import { Loading } from "VSTS_Extension/Components/Common/Loading";
import { BaseStore } from "VSTS_Extension/Flux/Stores/BaseStore";
import { MessagePanel, MessageType } from "VSTS_Extension/Components/Common/MessagePanel";
import { Hub, FilterPosition } from "VSTS_Extension/Components/Common/Hub/Hub";

import { CommandButton } from "OfficeFabric/Button";
import { autobind } from "OfficeFabric/Utilities";
import { Overlay } from "OfficeFabric/Overlay";
import { IContextualMenuItem } from "OfficeFabric/components/ContextualMenu/ContextualMenu.Props";

import { IBugBash, IBugBashViewModel } from "../Interfaces";
import { StoresHub } from "../Stores/StoresHub";
import { confirmAction, BugBashHelpers } from "../Helpers";
import { BugBashActions } from "../Actions/BugBashActions";
import { BugBashItemActions } from "../Actions/BugBashItemActions";
import { UrlActions } from "../Constants";

export interface IBugBashViewProps extends IBaseComponentProps {
    bugBashId?: string;
    pivotKey: string;
}

export interface IBugBashViewState extends IBaseComponentState {
    bugBashViewModel: IBugBashViewModel;
    loading: boolean;
    selectedPivot?: string;
    isAnyBugBashItemDirty?: boolean;
    bugBashEditorError?: string;
    filterText?: string;
}

export class BugBashView extends BaseComponent<IBugBashViewProps, IBugBashViewState> {
    protected initializeState() {
        this.state = {
            bugBashViewModel: this.props.bugBashId ? null : BugBashHelpers.getNewViewModel(),
            loading: this.props.bugBashId ? true : false,
            selectedPivot: this.props.pivotKey
        };
    }

    protected getStores(): BaseStore<any, any, any>[] {
        return [StoresHub.bugBashStore];
    }

    protected getStoresState(): IBugBashViewState {
        return {
            loading: StoresHub.bugBashStore.isLoading(this.props.bugBashId),
            bugBashEditorError: null,
            bugBashViewModel: BugBashHelpers.getViewModel(StoresHub.bugBashStore.getItem(this.props.bugBashId))
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
            let bugBash: IBugBash = null;

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
            // no-op
        }
    }

    public render(): JSX.Element {
        if (!this.state.loading && !this.state.bugBashViewModel) {
            return <MessagePanel messageType={MessageType.Error} message="This instance of bug bash doesn't exist." />;
        }
        else if(!this.state.loading && this.state.bugBashViewModel && !Utils_String.equals(VSS.getWebContext().project.id, this.state.bugBashViewModel.originalBugBash.projectId, true)) {
            return <MessagePanel messageType={MessageType.Error} message="This instance of bug bash is out of scope of current project." />;
        }
        else {
            return <div className="bugbash-view">
                { this.state.loading && <Overlay className="loading-overlay"><Loading /></Overlay> }
                { this.state.bugBashViewModel && this._renderHub() }
            </div>;
        }            
    }

    private _onTitleRender(title: string): React.ReactNode {
        return (
            <div className="bugbash-hub-title">
                <CommandButton 
                    text="Bug Bashes" 
                    onClick={async () => {
                        const confirm = await confirmAction(BugBashHelpers.isDirty(this.state.bugBashViewModel), "You have unsaved changes in the bug bash. Navigating to Home will revert your changes. Are you sure you want to do that?");

                        if (confirm) {
                            let navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
                            navigationService.updateHistoryEntry(UrlActions.ACTION_ALL, null);
                        } 
                    }}
                    className="bugbashes-button" />
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
                        let content: JSX.Element = null;
                        switch (key) {
                            case "edit":
                                content = this._renderEditor();
                                break;
                            case "results":
                                content = this._renderResults();
                                break;
                            case "charts":
                                content = this._renderCharts();
                                break;
                        }

                        return <div className="bugbash-hub-contents">
                            {content}
                        </div>;
                    },
                    pivots: [
                        {
                            key: "results",
                            text: "Results",
                            commands: this._getResultViewCommands(),
                            filterProps: {
                                showFilter: true,
                                onFilterChange: (filterText) => {
                                    this.updateState({filterText: filterText} as IBugBashViewState);
                                },
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
                            commands: this._getResultViewCommands(),
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
                    onChange={updatedBugbash => {
                        let newViewModel = {...this.state.bugBashViewModel};
                        newViewModel.updatedBugBash = {...updatedBugbash};
                        this.updateState({
                            bugBashViewModel: newViewModel
                        } as IBugBashViewState);
                    }}
                    save={() => this._saveBugBash()}
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
                        bugBash={{...this.state.bugBashViewModel.originalBugBash}} />
                )}
            </LazyLoad>;
        }
        else {
            return <MessagePanel messageType={MessageType.Info} message="Please save the bug bash first to view results" />;
        }
    }

    private _renderCharts(): JSX.Element {
        if (!BugBashHelpers.isNew(this.state.bugBashViewModel.originalBugBash)) {
            return <LazyLoad module="scripts/BugBashCharts">
                {(BugBashCharts) => (
                    <BugBashCharts.BugBashCharts bugBash={{...this.state.bugBashViewModel.originalBugBash}} />
                )}
            </LazyLoad>;
        }
        else {
            return <MessagePanel messageType={MessageType.Info} message="Please save the bug bash first to view charts" />;
        }
    }

    private _getEditorViewCommands(): IContextualMenuItem[] {
        return [
            {
                key: "save", name: "Save", iconProps: {iconName: "Save"}, 
                disabled: !BugBashHelpers.isDirty(this.state.bugBashViewModel) || !BugBashHelpers.isValid(this.state.bugBashViewModel),
                onClick: () => {
                    this._saveBugBash();
                }
            },
            {
                key: "undo", name: "Undo", iconProps: {iconName: "Undo"}, 
                disabled: BugBashHelpers.isNew(this.state.bugBashViewModel.originalBugBash) || !BugBashHelpers.isDirty(this.state.bugBashViewModel),
                onClick: async () => {
                    const confirm = await confirmAction(true, "Are you sure you want to undo your changes to this instance?");
                    if (confirm) {
                        let newViewModel = {...this.state.bugBashViewModel};
                        newViewModel.updatedBugBash = {...newViewModel.originalBugBash};
                        this.updateState({
                            bugBashViewModel: newViewModel
                        } as IBugBashViewState);
                    }
                }
            },
            {
                key: "refresh", name: "Refresh", iconProps: {iconName: "Refresh"}, 
                disabled: BugBashHelpers.isNew(this.state.bugBashViewModel.originalBugBash),
                onClick: async () => {
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
            }
        ];
    }

    private _getResultViewCommands(): IContextualMenuItem[] {
        return [
            {
                key: "refresh", name: "Refresh", iconProps: {iconName: "Refresh"}, 
                disabled: BugBashHelpers.isNew(this.state.bugBashViewModel.originalBugBash),
                onClick: async () => {
                    const confirm = await confirmAction(this.state.isAnyBugBashItemDirty, "You have some unsaved items in the list. Refreshing the page will remove all the unsaved data. Are you sure you want to do it?");
                    if (confirm) {
                        BugBashItemActions.refreshItems(this.props.bugBashId);
                    }
                }
            },
            {
                key: "newitem", name: "New", iconProps: {iconName: "Add"}, 
                disabled: BugBashHelpers.isNew(this.state.bugBashViewModel.originalBugBash),
                onClick: () => {
                    
                }
            }
        ];
    }

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