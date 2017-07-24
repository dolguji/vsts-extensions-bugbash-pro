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

import { IBugBash } from "../Interfaces";
import { StoresHub } from "../Stores/StoresHub";
import { confirmAction, BugBashHelpers } from "../Helpers";
import { BugBashActions } from "../Actions/BugBashActions";
import { BugBashItemActions } from "../Actions/BugBashItemActions";
import { UrlActions } from "../Constants";
import { BugBashProvider } from "../BugBashProvider";
import { BugBashItemProvider } from "../BugBashItemProvider";

export interface IBugBashViewProps extends IBaseComponentProps {    
    bugBashId?: string;
    pivotKey: string;
}

export interface IBugBashViewState extends IBaseComponentState {
    bugBash: IBugBash;
    loading: boolean;
    selectedPivot?: string;
    isBugBashDirty?: boolean;
    isBugBashValid?: boolean;
    isAnyBugBashItemDirty?: boolean;
    bugBashEditorError?: string;
    filterText?: string;
}

export class BugBashView extends BaseComponent<IBugBashViewProps, IBugBashViewState> {
    private _bugBashProvider: BugBashProvider;
    private _bugBashItemProvider: BugBashItemProvider;

    protected initializeState() {
        this._bugBashProvider = new BugBashProvider();
        this._bugBashItemProvider = new BugBashItemProvider();

        this.state = {
            bugBash: null,
            loading: true,
            selectedPivot: this.props.pivotKey
        };
    }

    protected getStores(): BaseStore<any, any, any>[] {
        return [StoresHub.bugBashStore];
    }

    protected getStoresState(): IBugBashViewState {
        if (this.props.bugBashId) {
            const bugBash = StoresHub.bugBashStore.getItem(this.props.bugBashId);
            const loading = StoresHub.bugBashStore.isLoading(this.props.bugBashId);
            let newState = {
                loading: loading,
                bugBashEditorError: null
            } as IBugBashViewState;
            
            if (!loading) {
                if (bugBash) {
                    this._bugBashProvider.initialize(bugBash);
                }
                newState = {...newState, bugBash: bugBash && {...bugBash}};
            }

            return newState;
        }
        else {
            return {
                loading: StoresHub.bugBashStore.isLoading()
            } as IBugBashViewState;
        }
    }

    public componentDidMount(): void {
        super.componentDidMount();

        this._bugBashProvider.addChangedListener(this._onProviderChange);
        this._bugBashItemProvider.addChangedListener(this._onProviderChange);

        let bugBash: IBugBash = null;
        if (!this.props.bugBashId) {
            bugBash = BugBashHelpers.getNewBugBash();
            this._bugBashProvider.initialize(bugBash);
            this.updateState({
                bugBash: bugBash,
                loading: false
            } as IBugBashViewState);
        }
        else {
            this._initializeBugBash(this.props.bugBashId);
        }
    }  

    public componentWillUnmount() {
        super.componentWillUnmount();

        this._bugBashProvider.removeChangedListener(this._onProviderChange);
        this._bugBashItemProvider.removeChangedListener(this._onProviderChange);
        this._bugBashProvider.dispose();
        this._bugBashItemProvider.dispose();
    }
    
    public componentWillReceiveProps(nextProps: Readonly<IBugBashViewProps>): void {
        if (nextProps.bugBashId !== this.props.bugBashId) {
            let bugBash: IBugBash = null;

            if (!nextProps.bugBashId) {
                bugBash = BugBashHelpers.getNewBugBash();
            }
            else {
                bugBash = StoresHub.bugBashStore.getItem(nextProps.bugBashId);
            } 
            
            this.updateState({
                bugBash: bugBash && {...bugBash},
                loading: bugBash == null ? true : false,
                selectedPivot: nextProps.pivotKey
            });
            
            if (bugBash) {
                this._bugBashProvider.initialize(bugBash);                
            }
            else {
                this._initializeBugBash(nextProps.bugBashId);            
            }
        }
        else {
            this.updateState({
                selectedPivot: nextProps.pivotKey
            } as IBugBashViewState);
        }
    }

    @autobind
    private _onProviderChange() {
        this.updateState({
            isBugBashDirty: this._bugBashProvider.isDirty(),
            isBugBashValid: this._bugBashProvider.isValid(),
            isAnyBugBashItemDirty: this._bugBashItemProvider.isAnyItemDirty(),
            bugBash: this._bugBashProvider.updatedBugBash
        } as IBugBashViewState);
    }

    private async _initializeBugBash(bugBashId: string) {        
        try {
            await BugBashActions.initializeBugBash(bugBashId);
            await BugBashItemActions.refreshItems(bugBashId);
        }
        catch (e) {
            // no-op
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
            return <div className="bugbash-view">
                { this.state.loading && <Overlay className="loading-overlay"><Loading /></Overlay> }
                { this._renderHub() }
            </div>;
        }            
    }

    private _onTitleRender(title: string): React.ReactNode {
        return (
            <div className="bugbash-hub-title">
                <CommandButton 
                    text="Bug Bashes" 
                    onClick={async () => {
                        const confirm = await confirmAction(this.state.isBugBashDirty, "You have unsaved changes in the bug bash. Navigating to Home will revert your changes. Are you sure you want to do that?");

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
        if (this.state.bugBash) {
            let title = this.state.bugBash.title;
            let className = "bugbash-hub";
            if (this.state.isBugBashDirty) {
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
                    bugBash={this.state.bugBash}
                    error={this.state.bugBashEditorError} 
                    provider={this._bugBashProvider}
                    save={() => this._saveBugBash()}
                />
            )}
        </LazyLoad>;
    }
    
    private _renderResults(): JSX.Element {
        if (!BugBashHelpers.isNew(this.state.bugBash)) {
            return <LazyLoad module="scripts/BugBashResults">
                {(BugBashResults) => (
                    <BugBashResults.BugBashResults 
                        filterText={this.state.filterText || ""}
                        bugBash={StoresHub.bugBashStore.getItem(this.state.bugBash.id)}
                        provider={this._bugBashItemProvider} />
                )}
            </LazyLoad>;
        }
        else {
            return <MessagePanel messageType={MessageType.Info} message="Please save the bug bash first to view results" />;
        }
    }

    private _renderCharts(): JSX.Element {
        if (!BugBashHelpers.isNew(this.state.bugBash)) {
            return <LazyLoad module="scripts/BugBashCharts">
                {(BugBashCharts) => (
                    <BugBashCharts.BugBashCharts bugBash={StoresHub.bugBashStore.getItem(this.state.bugBash.id)} />
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
                key: "save", name: "Save", iconProps: {iconName: "Save"}, disabled: !this.state.isBugBashDirty || !this.state.isBugBashValid,
                onClick: () => {
                    this._saveBugBash();
                }
            },
            {
                key: "undo", name: "Undo", iconProps: {iconName: "Undo"}, disabled: BugBashHelpers.isNew(this.state.bugBash) || !this.state.isBugBashDirty,
                onClick: async () => {
                    const confirm = await confirmAction(true, "Are you sure you want to undo your changes to this instance?");
                    if (confirm) {
                        this._bugBashProvider.undo();
                    }
                }
            },
            {
                key: "refresh", name: "Refresh", iconProps: {iconName: "Refresh"}, disabled: BugBashHelpers.isNew(this.state.bugBash),
                onClick: async () => {
                    const confirm = await confirmAction(this.state.isBugBashDirty, "Refreshing the item will undo your unsaved changes. Are you sure you want to do that?");
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
                key: "refresh", name: "Refresh", iconProps: {iconName: "Refresh"}, disabled: BugBashHelpers.isNew(this.state.bugBash),
                onClick: async () => {
                    const confirm = await confirmAction(this.state.isAnyBugBashItemDirty, "You have some unsaved items in the list. Refreshing the page will remove all the unsaved data. Are you sure you want to do it?");
                    if (confirm) {
                        BugBashItemActions.refreshItems(this.props.bugBashId);
                    }
                }
            },
            {
                key: "newitem", name: "New", iconProps: {iconName: "Add"}, disabled: BugBashHelpers.isNew(this.state.bugBash),
                onClick: () => {
                    
                }
            }
        ];
    }

    private async _saveBugBash() {
        if (BugBashHelpers.isNew(this.state.bugBash) && this.state.isBugBashDirty && this.state.isBugBashValid) {
            try {
                const createdBugBash = await BugBashActions.createBugBash(this.state.bugBash);
                let navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
                navigationService.updateHistoryEntry(UrlActions.ACTION_EDIT, { id: createdBugBash.id }, true);
            }
            catch (e) {
                this.updateState({bugBashEditorError: e} as IBugBashViewState);
            }
        }
        else if (!BugBashHelpers.isNew(this.state.bugBash) && this.state.isBugBashDirty && this.state.isBugBashValid) {
            try {
                await BugBashActions.updateBugBash(this.state.bugBash);
            }
            catch (e) {
                this.updateState({bugBashEditorError: e} as IBugBashViewState);
            }
        }
    }
}