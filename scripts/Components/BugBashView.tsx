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

import { Overlay } from "OfficeFabric/Overlay";
import { IContextualMenuItem } from "OfficeFabric/components/ContextualMenu/ContextualMenu.Props";

import { IBugBash } from "../Interfaces";
import { StoresHub } from "../Stores/StoresHub";
import { confirmAction, BugBashHelpers } from "../Helpers";
import { BugBashActions } from "../Actions/BugBashActions";
import { UrlActions } from "../Constants";

export interface IBugBashViewProps extends IBaseComponentProps {    
    bugBashId?: string;
    pivotKey: string;
}

export interface IBugBashViewState extends IBaseComponentState {
    originalBugBash: IBugBash;
    updatedBugBash: IBugBash;
    loading: boolean;
    selectedPivot?: string;
    bugBashEditorError?: string;
}

export class BugBashView extends BaseComponent<IBugBashViewProps, IBugBashViewState> {
    protected initializeState() {
        let bugBash = null;
        if (!this.props.bugBashId) {
            bugBash = BugBashHelpers.getNewBugBash();
        }
        
        this.state = {
            updatedBugBash: bugBash && {...bugBash},
            originalBugBash: bugBash && {...bugBash},
            loading: this.props.bugBashId ? true : false,
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
                newState = {...newState, updatedBugBash: bugBash && {...bugBash}, originalBugBash: bugBash && {...bugBash}};
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
        if (this.props.bugBashId) {
            this._initializeBugBash(this.props.bugBashId);
        }        
    }  

    public componentWillReceiveProps(nextProps: Readonly<IBugBashViewProps>): void {
        let bugBash = null;
        if (!nextProps.bugBashId) {
            bugBash = BugBashHelpers.getNewBugBash();
        } 
        else {
            bugBash = StoresHub.bugBashStore.getItem(nextProps.bugBashId);
        }          

        if (nextProps.bugBashId !== this.props.bugBashId) {
            this.updateState({
                updatedBugBash: bugBash && {...bugBash},
                originalBugBash: bugBash && {...bugBash},
                loading: bugBash == null ? true : false,
                selectedPivot: nextProps.pivotKey
            });
            
            if (bugBash == null && nextProps.bugBashId) {
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
        if (!this.state.loading && !this.state.originalBugBash) {
            return <MessagePanel messageType={MessageType.Error} message="This instance of bug bash doesn't exist." />;
        }
        else if(!this.state.loading && this.state.originalBugBash && !Utils_String.equals(VSS.getWebContext().project.id, this.state.originalBugBash.projectId, true)) {
            return <MessagePanel messageType={MessageType.Error} message="This instance of bug bash is out of scope of current project." />;
        }
        else {
            return <div className="bugbash-view">
                { this.state.loading && <Overlay className="loading-overlay"><Loading /></Overlay> }
                { this._renderHub() }
            </div>;
        }            
    }

    private _renderHub(): React.ReactNode {
        if (this.state.originalBugBash) {
            let title = this.state.updatedBugBash.title;
            let className = "bugbash-hub";
            if (BugBashHelpers.isDirty(this.state.updatedBugBash, this.state.originalBugBash)) {
                className += " is-dirty";
                title = "* " + title;
            }

            return <Hub 
                className={className}
                title={title}
                pivotProps={{
                    onPivotClick: async (selectedPivotKey: string, ev?: React.MouseEvent<HTMLElement>) => {
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
                            farCommands: this._getFarCommands(),
                            filterProps: {
                                showFilter: true,
                                onFilterChange: (filterText) => console.log(filterText),
                                filterPosition: FilterPosition.Left
                            }
                        },
                        {
                            key: "edit",
                            text: "Editor",
                            commands: this._getEditorViewCommands(),
                            farCommands: this._getFarCommands()
                        },
                        {
                            key: "charts",
                            text: "Charts",
                            commands: this._getResultViewCommands(),
                            farCommands: this._getFarCommands()
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
                    bugBash={this.state.updatedBugBash}
                    error={this.state.bugBashEditorError} 
                    save={() => this._saveBugBash()}
                    onChange={(updatedBugBash: IBugBash) => {
                        this.updateState({updatedBugBash: updatedBugBash} as IBugBashViewState);
                    }} />
            )}
        </LazyLoad>;
    }
    
    private _renderResults(): JSX.Element {
        if (this.state.originalBugBash.id) {
            return <LazyLoad module="scripts/BugBashResults">
                {(BugBashResults) => (
                    <BugBashResults.BugBashResults bugBash={this.state.originalBugBash} />
                )}
            </LazyLoad>;
        }
        else {
            return <MessagePanel messageType={MessageType.Info} message="Please save the bug bash first to view results" />;
        }
    }

    private _renderCharts(): JSX.Element {
        if (this.state.originalBugBash.id) {
            return <LazyLoad module="scripts/BugBashCharts">
                {(BugBashCharts) => (
                    <BugBashCharts.BugBashCharts bugBash={this.state.originalBugBash} />
                )}
            </LazyLoad>;
        }
        else {
            return <MessagePanel messageType={MessageType.Info} message="Please save the bug bash first to view charts" />;
        }
    }

    private _getFarCommands(): IContextualMenuItem[] {
        return [{
            key: "Home", name: "Home", iconProps: {iconName: "Home"}, 
            onClick: async (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                const confirm = await confirmAction(
                    BugBashHelpers.isDirty(this.state.updatedBugBash, this.state.originalBugBash), 
                    "You have unsaved changes in the bug bash. Navigating to Home will revert your changes. Are you sure you want to do that?");

                if (confirm) {
                    let navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
                    navigationService.updateHistoryEntry(UrlActions.ACTION_ALL, null);
                }                
            }
        }];
    }

    private _getEditorViewCommands(): IContextualMenuItem[] {
        return [
            {
                key: "save", name: "Save", iconProps: {iconName: "Save"}, disabled: !BugBashHelpers.isDirty(this.state.updatedBugBash, this.state.originalBugBash) || !BugBashHelpers.isValid(this.state.updatedBugBash),
                onClick: (event?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem) => {
                    this._saveBugBash();
                }
            },
            {
                key: "undo", name: "Undo", iconProps: {iconName: "Undo"}, disabled: BugBashHelpers.isNew(this.state.originalBugBash) || !BugBashHelpers.isDirty(this.state.updatedBugBash, this.state.originalBugBash),
                onClick: async (event?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem) => {
                    const confirm = await confirmAction(true, "Are you sure you want to undo your changes to this instance?");
                    if (confirm) {
                        this.updateState({updatedBugBash: {...this.state.originalBugBash}} as IBugBashViewState);
                    }
                }
            },
            {
                key: "refresh", name: "Refresh", iconProps: {iconName: "Refresh"}, disabled: BugBashHelpers.isNew(this.state.originalBugBash),
                onClick: async (event?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem) => {
                    const confirm = await confirmAction(BugBashHelpers.isDirty(this.state.updatedBugBash, this.state.originalBugBash), "Refreshing the item will undo your unsaved changes. Are you sure you want to do that?");
                    if (confirm) {
                        try {
                            await BugBashActions.refreshBugBash(this.state.originalBugBash.id);
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
                key: "refresh", name: "Refresh", iconProps: {iconName: "Refresh"}, disabled: BugBashHelpers.isNew(this.state.originalBugBash),
                onClick: async (event?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem) => {
                    // const confirm = await confirmAction(this._isAnyBugBashItemViewModelDirty(), "You have some unsaved items in the list. Refreshing the page will remove all the unsaved data. Are you sure you want to do it?");
                    // if (confirm) {
                    //     BugBashItemActions.refreshItems(this.props.originalBugBash.id);
                    // }
                }
            }
        ];
    }

    private async _saveBugBash() {
        if (BugBashHelpers.isNew(this.state.originalBugBash) && BugBashHelpers.isDirty(this.state.updatedBugBash, this.state.originalBugBash) && BugBashHelpers.isValid(this.state.updatedBugBash)) {
            try {
                const createdBugBash = await BugBashActions.createBugBash(this.state.updatedBugBash);
                let navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
                navigationService.updateHistoryEntry(UrlActions.ACTION_EDIT, { id: createdBugBash.id }, true);
            }
            catch (e) {
                this.updateState({bugBashEditorError: e} as IBugBashViewState);
            }
        }
        else if (!BugBashHelpers.isNew(this.state.originalBugBash) && BugBashHelpers.isDirty(this.state.updatedBugBash, this.state.originalBugBash) && BugBashHelpers.isValid(this.state.updatedBugBash)) {
            try {
                await BugBashActions.updateBugBash(this.state.updatedBugBash);
            }
            catch (e) {
                this.updateState({bugBashEditorError: e} as IBugBashViewState);
            }
        }
    }
}