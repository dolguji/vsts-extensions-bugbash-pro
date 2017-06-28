import "../../css/BugBashResultsView.scss";

import * as React from "react";

import { HostNavigationService } from "VSS/SDK/Services/Navigation";
import Utils_Date = require("VSS/Utils/Date");
import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");
import { WorkItem } from "TFS/WorkItemTracking/Contracts";
import * as WitClient from "TFS/WorkItemTracking/RestClient";

import { PrimaryButton } from "OfficeFabric/Button";
import { Label } from "OfficeFabric/Label";
import { IContextualMenuItem } from "OfficeFabric/components/ContextualMenu/ContextualMenu.Props";
import { autobind } from "OfficeFabric/Utilities";
import { Pivot, PivotItem } from "OfficeFabric/Pivot";
import { TooltipHost, TooltipDelay, DirectionalHint, TooltipOverflowMode } from "OfficeFabric/Tooltip";
import { SelectionMode } from "OfficeFabric/utilities/selection/interfaces";
import { CommandBar } from "OfficeFabric/CommandBar";

import { MessagePanel, MessageType } from "VSTS_Extension/Components/Common/MessagePanel";
import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { BaseStore } from "VSTS_Extension/Stores/BaseStore";
import { WorkItemFieldStore } from "VSTS_Extension/Stores/WorkItemFieldStore";
import { Loading } from "VSTS_Extension/Components/Common/Loading";
import { Grid } from "VSTS_Extension/Components/Grids/Grid";
import { WorkItemGrid } from "VSTS_Extension/Components/Grids/WorkItemGrid/WorkItemGrid";
import { IdentityView } from "VSTS_Extension/Components/WorkItemControls/IdentityView";
import { SortOrder, GridColumn } from "VSTS_Extension/Components/Grids/Grid.Props";
import { ColumnPosition, IExtraWorkItemGridColumn } from "VSTS_Extension/Components/Grids/WorkItemGrid/WorkItemGrid.Props";
import { LazyLoad } from "VSTS_Extension/Components/Common/LazyLoad";

import { UrlActions } from "../Constants";
import { IBugBash, IBugBashItem, IBugBashItemViewModel } from "../Interfaces";
import { confirmAction, BugBashItemHelpers } from "../Helpers";
import { BugBashItemManager } from "../BugBashItemManager";
import { BugBashItemEditor } from "./BugBashItemEditor";
import { BugBashStore } from "../Stores/BugBashStore";
import { TeamStore } from "../Stores/TeamStore";
import { StoresHub } from "../Stores/StoresHub";


interface IBugBashResultsViewState extends IBaseComponentState {
    bugBashItem?: IBugBash;
    bugBashItemDoesntExist?: boolean;
    viewModels?: IBugBashItemViewModel[];
    selectedViewModel?: IBugBashItemViewModel;
    workItemsMap?: IDictionaryNumberTo<WorkItem>;
    selectedPivot?: SelectedPivot;
}

interface IBugBashResultsViewProps extends IBaseComponentProps {
    id: string;
}

enum SelectedPivot {
    Pending,
    Accepted,
    Rejected,
    Analytics
}

export class BugBashResultsView extends BaseComponent<IBugBashResultsViewProps, IBugBashResultsViewState> {
    private _itemInvokedTimeout: any;

    protected getStoresToLoad(): {new (): BaseStore<any, any, any>}[] {
        return [BugBashStore, WorkItemFieldStore, TeamStore];
    }

    protected initializeState() {
        this.state = {
            bugBashItem: null,
            bugBashItemDoesntExist: false,
            viewModels: null,
            selectedViewModel: null,
            workItemsMap: null,
            selectedPivot: SelectedPivot.Pending
        };
    }

    protected async initialize() {            
        const found = await StoresHub.bugBashStore.ensureItem(this.props.id);

        if (!found) {
            this.updateState({
                bugBashItem: null,
                viewModels: [],
                selectedViewModel: null,
                bugBashItemDoesntExist: true
            });
        }
        else {
            StoresHub.teamStore.initialize();
            StoresHub.workItemFieldStore.initialize();
            this._refreshData();
        }
    }   

    private async _refreshData() {
        this.updateState({workItemsMap: null, viewModels: null, selectedViewModel: null});

        let items = await BugBashItemManager.getItems(this.props.id);
        const workItemIds = items.filter(item => BugBashItemHelpers.isAccepted(item)).map(item => item.workItemId);

        if (workItemIds.length > 0) {
            let workItems = await WitClient.getClient().getWorkItems(workItemIds, ["System.Id", "System.Title", "System.WorkItemType", "System.State", "System.AssignedTo", "System.AreaPath"]);
            let map: IDictionaryNumberTo<WorkItem> = {};
            for (const workItem of workItems) {
                if (workItem) {
                    map[workItem.id] = workItem;
                }
            }
            this.updateState({workItemsMap: map, viewModels: items.map(item => BugBashItemHelpers.getItemViewModel(item))});
        }
        else {
            this.updateState({workItemsMap: {}, viewModels: items.map(item => BugBashItemHelpers.getItemViewModel(item))});
        }
    }

    protected onStoreChanged() {
        const bugBashItem = StoresHub.bugBashStore.getItem(this.props.id);

        this.updateState({
            bugBashItem: bugBashItem,
            selectedPivot: bugBashItem.autoAccept ? SelectedPivot.Accepted : SelectedPivot.Pending
        });
    }

    private _isDataLoading(): boolean {
        if (this.state.bugBashItemDoesntExist) {
            return false;
        }
        
        return !StoresHub.bugBashStore.isLoaded() 
            || this.state.workItemsMap == null 
            || !StoresHub.workItemFieldStore.isLoaded() 
            || !StoresHub.teamStore.isLoaded()
            || this.state.viewModels == null;
    }

    private _getSelectedPivotKey(): string {
        switch (this.state.selectedPivot) {
            case SelectedPivot.Pending:
                return "Pending";
            case SelectedPivot.Accepted:
                return "Accepted";
            case SelectedPivot.Rejected:
                return "Rejected";
            default:
                return "Analytics";
        }
    }

    private _updateSelectedPivot(pivotKey: string) {
        switch (pivotKey) {
            case "Pending":
                this.updateState({selectedPivot: SelectedPivot.Pending});
                break;
            case "Accepted":
                this.updateState({selectedPivot: SelectedPivot.Accepted});
                break;
            case "Rejected":
                this.updateState({selectedPivot: SelectedPivot.Rejected});
                break;
            default:
                this.updateState({selectedPivot: SelectedPivot.Analytics});
                break;
        }
    }
    
    public render(): JSX.Element {
        if (this._isDataLoading()) {
            return <Loading />;
        }
        else {
            if (!this.state.bugBashItem) {
                return <MessagePanel messageType={MessageType.Error} message="This instance of bug bash doesn't exist." />;
            }
            else if(!Utils_String.equals(VSS.getWebContext().project.id, this.state.bugBashItem.projectId, true)) {
                return <MessagePanel messageType={MessageType.Error} message="This instance of bug bash is out of scope of current project." />;
            }
            else {
                return (
                    <div className="results-view">
                        <Label className="bugbash-title overflow-ellipsis">{StoresHub.bugBashStore.getItem(this.props.id).title}</Label>
                        <div className="results-container">
                            <div className="left-content">
                                {this._renderPivots()}
                            </div>
                            <div className="right-content">
                                {this._renderItemEditor()}
                            </div>    
                        </div>
                    </div>
                );
            }
        }
    }

    private _renderPivots(): JSX.Element {
        const allViewModels = this.state.viewModels.filter(viewModel => !BugBashItemHelpers.isAccepted(viewModel.originalModel) || this.state.workItemsMap[viewModel.originalModel.workItemId] != null);
        const pendingItems = allViewModels.filter(viewModel => !BugBashItemHelpers.isAccepted(viewModel.originalModel) && !viewModel.originalModel.rejected);
        const rejectedItems = allViewModels.filter(viewModel => viewModel.originalModel.rejected);
        const workItemIds = allViewModels.filter(viewModel => BugBashItemHelpers.isAccepted(viewModel.originalModel)).map(viewModel => viewModel.originalModel.workItemId);

        let workItems: WorkItem[] = [];
        for (let id of workItemIds) {
            if (this.state.workItemsMap[id]) {
                workItems.push(this.state.workItemsMap[id]);
            }
        }

        const fields = [
            StoresHub.workItemFieldStore.getItem("System.Id"),
            StoresHub.workItemFieldStore.getItem("System.Title"),
            StoresHub.workItemFieldStore.getItem("System.State"),
            StoresHub.workItemFieldStore.getItem("System.AreaPath")
        ];
        
        let pivots: JSX.Element;
        let pivotContent: JSX.Element;

        if (this.state.bugBashItem.autoAccept) {
            pivots = <Pivot initialSelectedKey={this._getSelectedPivotKey()} onLinkClick={(item: PivotItem) => this._updateSelectedPivot(item.props.itemKey)}>
                <PivotItem linkText={`Accepted Items (${workItems.length})`} itemKey="Accepted" />
                <PivotItem linkText="Analytics" itemKey="Analytics" />
            </Pivot>;
        }
        else {
            pivots = <Pivot initialSelectedKey={this._getSelectedPivotKey()} onLinkClick={(item: PivotItem) => this._updateSelectedPivot(item.props.itemKey)}>
                <PivotItem linkText={`Pending Items (${pendingItems.length})`} itemKey="Pending" />
                <PivotItem linkText={`Accepted Items (${workItems.length})`} itemKey="Accepted" />
                <PivotItem linkText={`Rejected Items (${rejectedItems.length})`} itemKey="Rejected" />
                <PivotItem linkText="Analytics" itemKey="Analytics" />
            </Pivot>;
        }
        
        switch (this.state.selectedPivot) {
            case SelectedPivot.Pending:
                pivotContent = <Grid
                    selectionPreservedOnEmptyClick={true}
                    setKey="bugbash-pending-item-grid"
                    className="bugbash-item-grid"
                    items={pendingItems}
                    selectionMode={SelectionMode.single}
                    columns={this._getBugBashItemGridColumns(false)}                    
                    commandBarProps={{menuItems: this._getCommandBarMenuItems(), farMenuItems: this._getCommandBarFarMenuItems()}}
                    events={{
                        onSelectionChanged: this._onBugBashItemSelectionChanged
                    }}
                />;
                break;
            case SelectedPivot.Accepted:
                pivotContent = <WorkItemGrid
                    selectionPreservedOnEmptyClick={true}
                    setKey="bugbash-work-item-grid"
                    className="bugbash-item-grid"
                    workItems={workItems}                    
                    fields={fields}
                    noResultsText="No Accepted items"
                    extraColumns={this._getExtraWorkItemGridColumns()}
                    onWorkItemUpdated={(updatedWorkItem: WorkItem) => {
                        let map = {...this.state.workItemsMap};
                        map[updatedWorkItem.id] = updatedWorkItem;
                        this.updateState({workItemsMap: map});
                    }}
                    commandBarProps={{menuItems: this._getCommandBarMenuItems(), farMenuItems: this._getCommandBarFarMenuItems()}}
                />;
                break;
            case SelectedPivot.Rejected:
                pivotContent = <Grid
                    selectionPreservedOnEmptyClick={true}
                    setKey="bugbash-rejected-item-grid"
                    className="bugbash-item-grid"
                    items={rejectedItems}
                    selectionMode={SelectionMode.single}
                    columns={this._getBugBashItemGridColumns(true)}
                    commandBarProps={{menuItems: this._getCommandBarMenuItems(), farMenuItems: this._getCommandBarFarMenuItems()}}
                    events={{
                        onSelectionChanged: this._onBugBashItemSelectionChanged
                    }}
                />;
                break;
            default:
                pivotContent = (
                    <div className="analytics-container">
                        <CommandBar 
                            className="analytics-editor-menu"
                            items={this._getCommandBarMenuItems()} 
                            farItems={this._getCommandBarFarMenuItems()} />

                        <LazyLoad module="scripts/BugBashResultsAnalytics">
                            {(BugBashResultsAnalytics) => (
                                <BugBashResultsAnalytics.BugBashResultsAnalytics itemModels={allViewModels.map(vm => vm.originalModel)} />
                            )}
                        </LazyLoad>
                    </div>
                );                
                break;
        }

        return (
            <div className="pivot-container">
                {pivots}
                {pivotContent}
            </div>
        );
    }

    private _renderItemEditor(): JSX.Element {
        return (
            <div className="item-editor-container">
                <div className="item-editor-container-header">
                    <Label className="item-editor-header overflow-ellipsis">{this.state.selectedViewModel ? "Edit Item" : "Add Item"}</Label>
                    <PrimaryButton    
                        className="add-new-item-button"                    
                        disabled={this.state.selectedViewModel == null}
                        text="New"
                        iconProps={ { iconName: "Add" } }
                        onClick={ () => this.updateState({selectedViewModel: null}) }
                        />
                </div>
                <BugBashItemEditor 
                    viewModel={this.state.selectedViewModel || BugBashItemHelpers.getNewItemViewModel(this.props.id)}
                    onItemAccept={(model: IBugBashItem, workItem: WorkItem) => {
                        let newViewModels = this.state.viewModels.slice();
                        let newViewModel = BugBashItemHelpers.getItemViewModel(model);
                        let newWorkItemsMap = {...this.state.workItemsMap};

                        let index = Utils_Array.findIndex(newViewModels, v => v.model.id === model.id);
                        if (index !== -1) {
                            newViewModels[index] = newViewModel;
                        }
                        else {
                            newViewModels.push(newViewModel);
                        }

                        if (workItem) {
                            newWorkItemsMap[workItem.id] = workItem;
                            this.updateState({viewModels: newViewModels, selectedViewModel: null, workItemsMap: newWorkItemsMap});
                        }
                        else {
                            this.updateState({viewModels: newViewModels, selectedViewModel: newViewModel});
                        }
                    }}
                    onItemUpdate={(model: IBugBashItem) => {
                        if (model.id) {
                            let newViewModels = this.state.viewModels.slice();
                            let newVideModel = BugBashItemHelpers.getItemViewModel(model);

                            let index = Utils_Array.findIndex(newViewModels, v => v.model.id === model.id);
                            if (index !== -1) {
                                newViewModels[index] = newVideModel;
                            }
                            else {
                                newViewModels.push(newVideModel);
                            }
                            this.updateState({viewModels: newViewModels, selectedViewModel: newVideModel});
                        }
                    }}
                    onDelete={(model: IBugBashItem) => {
                        if (model.id) {
                            let newViewModels = this.state.viewModels.slice();
                            let index = Utils_Array.findIndex(newViewModels, v => v.model.id === model.id);
                            if (index !== -1) {
                                newViewModels.splice(index, 1);
                                this.updateState({viewModels: newViewModels, selectedViewModel: null});
                            }
                            else {
                                this.updateState({selectedViewModel: null});
                            }
                        }
                    }} 
                    onChange={(changedModel: IBugBashItem, newComment: string) => {
                        if (changedModel.id) {
                            let newViewModels = this.state.viewModels.slice();
                            let index = Utils_Array.findIndex(newViewModels, viewModel => viewModel.model.id === changedModel.id);
                            if (index !== -1) {
                                newViewModels[index].model = {...changedModel};
                                newViewModels[index].newComment = newComment;
                                this.updateState({viewModels: newViewModels, selectedViewModel: newViewModels[index]});
                            }
                        }
                    }} />
            </div>
        );
    }

    private _getExtraWorkItemGridColumns(): IExtraWorkItemGridColumn[] {
        let workItemIdToItemMap: IDictionaryNumberTo<IBugBashItem> = {};
        for (const viewModel of this.state.viewModels) {
            if (BugBashItemHelpers.isAccepted(viewModel.model)) {
                workItemIdToItemMap[viewModel.model.workItemId] = viewModel.model;
            }
        }

        return [
            {
                position: ColumnPosition.FarRight,
                column: {
                    key: "createdby",
                    name: "Item created by",
                    minWidth: 100,
                    maxWidth: 250,
                    resizable: true,
                    onRenderCell: (workItem: WorkItem) => {
                        return <IdentityView identityDistinctName={workItemIdToItemMap[workItem.id].createdBy} />;
                    },
                    sortFunction: (workItem1: WorkItem, workItem2: WorkItem, sortOrder: SortOrder) => {                        
                        let compareValue = Utils_String.ignoreCaseComparer(workItemIdToItemMap[workItem1.id].createdBy, workItemIdToItemMap[workItem2.id].createdBy);
                        return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                    },
                    filterFunction: (workItem: WorkItem, filterText: string) => Utils_String.caseInsensitiveContains(workItemIdToItemMap[workItem.id].createdBy, filterText)
                }
            }
        ];
    }
    
    @autobind
    private _onBugBashItemSelectionChanged(viewModels: IBugBashItemViewModel[]) {
        if (this._itemInvokedTimeout) {
            clearTimeout(this._itemInvokedTimeout);
            this._itemInvokedTimeout = null;
        }

        this._itemInvokedTimeout = setTimeout(() => {
            if (viewModels == null || viewModels.length !== 1) {
                this.updateState({selectedViewModel: null});
            }
            else {
                this.updateState({selectedViewModel: viewModels[0]});
            }

            this._itemInvokedTimeout = null;
        }, 200);
    }

    private _getBugBashItemGridColumns(isRejectedGrid: boolean): GridColumn[] {
        const gridCellClassName = "item-grid-cell";
        const getCellClassName = (viewModel: IBugBashItemViewModel) => {
            let className = gridCellClassName;
            if (BugBashItemHelpers.isDirty(viewModel)) {
                className += " is-dirty";
            }
            if (!BugBashItemHelpers.isValid(viewModel.model)) {
                className += " is-invalid";
            }

            return className;
        }                

        let columns: GridColumn[] = [
            {
                key: "title",
                name: "Title",
                minWidth: 200,
                maxWidth: isRejectedGrid ? 600 : 800,
                resizable: true,
                onRenderCell: (viewModel: IBugBashItemViewModel) => {                 
                    return (
                        <TooltipHost 
                            content={viewModel.model.title}
                            delay={TooltipDelay.medium}
                            overflowMode={TooltipOverflowMode.Parent}
                            directionalHint={DirectionalHint.bottomLeftEdge}>

                            <Label className={`${getCellClassName(viewModel)}`}>
                                {`${BugBashItemHelpers.isDirty(viewModel) ? "* " : ""}${viewModel.model.title}`}
                            </Label>
                        </TooltipHost>
                    )
                },
                sortFunction: (viewModel1: IBugBashItemViewModel, viewModel2: IBugBashItemViewModel, sortOrder: SortOrder) => {
                    let compareValue = Utils_String.ignoreCaseComparer(viewModel1.model.title, viewModel2.model.title);
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                },
                filterFunction: (viewModel: IBugBashItemViewModel, filterText: string) => Utils_String.caseInsensitiveContains(viewModel.model.title, filterText)
            },
            {
                key: "team",
                name: "Team",
                minWidth: isRejectedGrid ? 100 : 200,
                maxWidth: isRejectedGrid ? 200 : 300,
                resizable: true,
                onRenderCell: (viewModel: IBugBashItemViewModel) => {
                    const team = StoresHub.teamStore.getItem(viewModel.model.teamId);
                    return (
                        <TooltipHost 
                            content={team ? team.name : viewModel.model.teamId}
                            delay={TooltipDelay.medium}
                            overflowMode={TooltipOverflowMode.Parent}
                            directionalHint={DirectionalHint.bottomLeftEdge}>

                            <Label className={`${getCellClassName(viewModel)}`}>
                                {team ? team.name : viewModel.model.teamId}
                            </Label>
                        </TooltipHost>
                    )
                },
                sortFunction: (viewModel1: IBugBashItemViewModel, viewModel2: IBugBashItemViewModel, sortOrder: SortOrder) => {
                    const team1 = StoresHub.teamStore.getItem(viewModel1.model.teamId);
                    const team2 = StoresHub.teamStore.getItem(viewModel2.model.teamId);
                    const team1Name = team1 ? team1.name : viewModel1.model.teamId;
                    const team2Name = team2 ? team2.name : viewModel2.model.teamId;

                    let compareValue = Utils_String.ignoreCaseComparer(team1Name, team2Name);
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                },
                filterFunction: (viewModel: IBugBashItemViewModel, filterText: string) =>  {
                    const team = StoresHub.teamStore.getItem(viewModel.model.teamId);
                    return Utils_String.caseInsensitiveContains(team ? team.name : viewModel.model.teamId, filterText);
                }
            },
            {
                key: "createdby",
                name: "Created By",
                minWidth: isRejectedGrid ? 100 : 200,
                maxWidth: isRejectedGrid ? 200 : 300,
                resizable: true,
                onRenderCell: (viewModel: IBugBashItemViewModel) => {
                    return (
                        <TooltipHost 
                            content={viewModel.model.createdBy}
                            delay={TooltipDelay.medium}
                            directionalHint={DirectionalHint.bottomLeftEdge}>

                            <div className={getCellClassName(viewModel)}>
                                <IdentityView identityDistinctName={viewModel.model.createdBy} />
                            </div>
                        </TooltipHost>
                    )
                },
                sortFunction: (viewModel1: IBugBashItemViewModel, viewModel2: IBugBashItemViewModel, sortOrder: SortOrder) => {
                    let compareValue = Utils_String.ignoreCaseComparer(viewModel1.model.createdBy, viewModel2.model.createdBy);
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                },
                filterFunction: (viewModel: IBugBashItemViewModel, filterText: string) => Utils_String.caseInsensitiveContains(viewModel.model.createdBy, filterText)
            },
            {
                key: "createddate",
                name: "Created Date",
                minWidth: isRejectedGrid ? 100 : 200,
                maxWidth: isRejectedGrid ? 200 : 300,
                resizable: true,
                onRenderCell: (viewModel: IBugBashItemViewModel) => {
                    return (
                        <TooltipHost 
                            content={Utils_Date.format(viewModel.model.createdDate, "M/d/yyyy h:mm tt")}
                            delay={TooltipDelay.medium}
                            overflowMode={TooltipOverflowMode.Parent}
                            directionalHint={DirectionalHint.bottomLeftEdge}>

                            <Label className={`${getCellClassName(viewModel)}`}>
                                {Utils_Date.friendly(viewModel.model.createdDate)}
                            </Label>
                        </TooltipHost>
                    )
                },
                sortFunction: (viewModel1: IBugBashItemViewModel, viewModel2: IBugBashItemViewModel, sortOrder: SortOrder) => {
                    let compareValue = Utils_Date.defaultComparer(viewModel1.model.createdDate, viewModel2.model.createdDate);
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                }
            }
        ];

        if (isRejectedGrid) {
            columns.push({
                key: "rejectedby",
                name: "Rejected By",
                minWidth: 100,
                maxWidth: 200,
                resizable: true,
                onRenderCell: (viewModel: IBugBashItemViewModel) => {
                    return (
                        <TooltipHost 
                            content={viewModel.model.rejectedBy}
                            delay={TooltipDelay.medium}
                            directionalHint={DirectionalHint.bottomLeftEdge}>

                            <div className={getCellClassName(viewModel)}>
                                <IdentityView identityDistinctName={viewModel.model.rejectedBy} />
                            </div>
                        </TooltipHost>
                    )
                },
                sortFunction: (viewModel1: IBugBashItemViewModel, viewModel2: IBugBashItemViewModel, sortOrder: SortOrder) => {
                    let compareValue = Utils_String.ignoreCaseComparer(viewModel1.model.rejectedBy, viewModel2.model.rejectedBy);
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                },
                filterFunction: (viewModel: IBugBashItemViewModel, filterText: string) => Utils_String.caseInsensitiveContains(viewModel.model.rejectedBy, filterText)
            }, 
            {
                key: "rejectreason",
                name: "Reject Reason",
                minWidth: 200,
                maxWidth: 800,
                resizable: true,
                onRenderCell: (viewModel: IBugBashItemViewModel) => {                 
                    return (
                        <TooltipHost 
                            content={viewModel.model.rejectReason}
                            delay={TooltipDelay.medium}
                            overflowMode={TooltipOverflowMode.Parent}
                            directionalHint={DirectionalHint.bottomLeftEdge}>

                            <Label className={`${getCellClassName(viewModel)}`}>
                                {viewModel.model.rejectReason}
                            </Label>
                        </TooltipHost>
                    )
                },
                sortFunction: (viewModel1: IBugBashItemViewModel, viewModel2: IBugBashItemViewModel, sortOrder: SortOrder) => {
                    let compareValue = Utils_String.ignoreCaseComparer(viewModel1.model.rejectReason, viewModel2.model.rejectReason);
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                },
                filterFunction: (viewModel: IBugBashItemViewModel, filterText: string) => Utils_String.caseInsensitiveContains(viewModel.model.rejectReason, filterText)
            });
        }

        return columns;
    }

    private _getCommandBarMenuItems(): IContextualMenuItem[] {
        return [
                {
                    key: "edit", name: "Edit", title: "Edit", iconProps: {iconName: "Edit"},
                    onClick: async () => {
                        const confirm = await confirmAction(this._isAnyViewModelDirty(), "You have some unsaved items in the list. Navigating to a different page will remove all the unsaved data. Are you sure you want to do it?");
                        if (confirm) {
                            let navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
                            navigationService.updateHistoryEntry(UrlActions.ACTION_EDIT, {id: this.props.id});
                        }
                    }
                },
                {
                    key: "refresh", name: "Refresh", title: "Refresh list", iconProps: {iconName: "Refresh"},
                    onClick: async () => {
                        const confirm = await confirmAction(this._isAnyViewModelDirty(), "You have some unsaved items in the list. Refreshing the page will remove all the unsaved data. Are you sure you want to do it?");
                        if (confirm) {
                            this._refreshData();
                        }
                    }
                }
            ];
    }

    private _getCommandBarFarMenuItems(): IContextualMenuItem[] {
        let menuItems = [
                {
                    key: "Home", name: "Home", title: "Return to home view", iconProps: {iconName: "Home"}, 
                    onClick: async () => {
                        const confirm = await confirmAction(this._isAnyViewModelDirty(), "You have some unsaved items in the list. Navigating to a different page will remove all the unsaved data. Are you sure you want to do it?");
                        if (confirm) {
                            let navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
                            navigationService.updateHistoryEntry(UrlActions.ACTION_ALL, null);
                        }
                    }
                }
            ] as IContextualMenuItem[];

        return menuItems;
    }

    private _isAnyViewModelDirty(viewModels?: IBugBashItemViewModel[]): boolean {
        for (let viewModel of viewModels || this.state.viewModels) {
            if (BugBashItemHelpers.isDirty(viewModel)) {
                return true;
            }
        }

        return false;
    }
}