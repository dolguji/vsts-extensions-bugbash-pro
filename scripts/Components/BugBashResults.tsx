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
import { BaseStore } from "VSTS_Extension/Flux/Stores/BaseStore";
import { WorkItemFieldStore } from "VSTS_Extension/Flux/Stores/WorkItemFieldStore";
import { TeamStore } from "VSTS_Extension/Flux/Stores/TeamStore";
import { WorkItemFieldActions } from "VSTS_Extension/Flux/Actions/WorkItemFieldActions";
import { TeamActions } from "VSTS_Extension/Flux/Actions/TeamActions";
import { Loading } from "VSTS_Extension/Components/Common/Loading";
import { Grid } from "VSTS_Extension/Components/Grids/Grid";
import { WorkItemGrid } from "VSTS_Extension/Components/Grids/WorkItemGrid/WorkItemGrid";
import { IdentityView } from "VSTS_Extension/Components/WorkItemControls/IdentityView";
import { SortOrder, GridColumn } from "VSTS_Extension/Components/Grids/Grid.Props";
import { ColumnPosition, IExtraWorkItemGridColumn } from "VSTS_Extension/Components/Grids/WorkItemGrid/WorkItemGrid.Props";
import { LazyLoad } from "VSTS_Extension/Components/Common/LazyLoad";
import SplitterLayout from "rc-split-layout";

import { UrlActions } from "../Constants";
import { IBugBash, IBugBashItem, IBugBashItemViewModel } from "../Interfaces";
import { confirmAction, BugBashItemHelpers } from "../Helpers";
import { BugBashItemEditor } from "./BugBashItemEditor";
import { StoresHub } from "../Stores/StoresHub";
import { BugBashActions } from "../Actions/BugBashActions";
import { BugBashItemActions } from "../Actions/BugBashItemActions";

interface IBugBashResultsState extends IBaseComponentState {
    bugBashItemViewModels: IBugBashItemViewModel[];
    workItemsMap: IDictionaryNumberTo<WorkItem>;
    selectedPivot: SelectedPivot;
    selectedBugBashItemViewModel?: IBugBashItemViewModel;
    loading?: boolean;
}

interface IBugBashResultsProps extends IBaseComponentProps {
    bugBash: IBugBash;
}

enum SelectedPivot {
    Pending,
    Accepted,
    Rejected,
    Analytics
}

export class BugBashResults extends BaseComponent<IBugBashResultsProps, IBugBashResultsState> {
    private _itemInvokedTimeout: any;

    protected initializeState() {
        this.state = {
            bugBashItemViewModels: null,
            selectedBugBashItemViewModel: null,
            workItemsMap: null,
            selectedPivot: SelectedPivot.Pending,
            loading: true
        };
    }

    protected getStores(): BaseStore<any, any, any>[] {
        return [StoresHub.bugBashItemStore];
    }

    protected getStoresState(): IBugBashResultsState {
        const bugBashItems = StoresHub.bugBashItemStore.getBugBashItems(this.props.bugBash.id);
        let workItemIds: number[];

        if (bugBashItems) {
            workItemIds = bugBashItems.filter(bugBashItem => BugBashItemHelpers.isAccepted(bugBashItem)).map(item => item.workItemId);
        }

        return {
            bugBashItemViewModels: bugBashItems ? bugBashItems.map(bugBashItem => BugBashItemHelpers.getBugBashItemViewModel(bugBashItem)) : null,
            loading: StoresHub.bugBashItemStore.isLoading(this.props.bugBash.id)
        } as IBugBashResultsState;
    }

    public componentDidMount() {
        super.componentDidMount();
        BugBashItemActions.initializeItems(this.props.bugBash.id);
    }

    public componentWillReceiveProps(nextProps: Readonly<IBugBashResultsProps>): void {
        if (nextProps.bugBash.id !== this.props.bugBash.id || nextProps.bugBash.__etag !== this.props.bugBash.__etag) {
            this.updateState({
                model: {...nextProps.bugBash},
                error: null
            });
        }
        else {
            this.updateState({
                error: null
            } as IBugBashResultsProps);
        }
    }

    private async _refreshData() {
        this.updateState({workItemsMap: null, bugBashItemViewModels: null, selectedBugBashItemViewModel: null});

        let items = [] //await BugBashItemManager.getItems(this.props.id);
        const workItemIds = items.filter(item => BugBashItemHelpers.isAccepted(item)).map(item => item.workItemId);

        if (workItemIds.length > 0) {
            let workItems = await WitClient.getClient().getWorkItems(workItemIds, ["System.Id", "System.Title", "System.WorkItemType", "System.State", "System.AssignedTo", "System.AreaPath"]);
            let map: IDictionaryNumberTo<WorkItem> = {};
            for (const workItem of workItems) {
                if (workItem) {
                    map[workItem.id] = workItem;
                }
            }
            this.updateState({workItemsMap: map, bugBashItemViewModels: items.map(item => BugBashItemHelpers.getBugBashItemViewModel(item))});
        }
        else {
            this.updateState({workItemsMap: {}, bugBashItemViewModels: items.map(item => BugBashItemHelpers.getBugBashItemViewModel(item))});
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
            || this.state.bugBashItemViewModels == null;
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
                            <SplitterLayout 
                                primaryIndex={0}
                                primaryMinSize={550}
                                secondaryMinSize={400}
                                secondaryInitialSize={500}
                                onChange={() => {
                                    let evt = document.createEvent('UIEvents');
                                    evt.initUIEvent('resize', true, false, window, 0);
                                    window.dispatchEvent(evt);
                                }} >
                                
                                <div className="left-content">
                                    {this._renderPivots()}
                                </div>
                                <div className="right-content">
                                    {this._renderItemEditor()}
                                </div>
                            </SplitterLayout>
                        </div>
                    </div>
                );
            }
        }
    }

    private _renderPivots(): JSX.Element {
        const allViewModels = this.state.bugBashItemViewModels.filter(viewModel => !BugBashItemHelpers.isAccepted(viewModel.originalBugBashItem) || this.state.workItemsMap[viewModel.originalBugBashItem.workItemId] != null);
        const pendingItems = allViewModels.filter(viewModel => !BugBashItemHelpers.isAccepted(viewModel.originalBugBashItem) && !viewModel.originalBugBashItem.rejected);
        const rejectedItems = allViewModels.filter(viewModel => viewModel.originalBugBashItem.rejected);
        const workItemIds = allViewModels.filter(viewModel => BugBashItemHelpers.isAccepted(viewModel.originalBugBashItem)).map(viewModel => viewModel.originalBugBashItem.workItemId);

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
                                <BugBashResultsAnalytics.BugBashResultsAnalytics itemModels={allViewModels.map(vm => vm.originalBugBashItem)} />
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
                    <Label className="item-editor-header overflow-ellipsis">{this.state.selectedBugBashItemViewModel ? "Edit Item" : "Add Item"}</Label>
                    <PrimaryButton    
                        className="add-new-item-button"                    
                        disabled={this.state.selectedBugBashItemViewModel == null}
                        text="New"
                        iconProps={ { iconName: "Add" } }
                        onClick={ () => this.updateState({selectedBugBashItemViewModel: null}) }
                        />
                </div>
                <BugBashItemEditor 
                    viewModel={this.state.selectedBugBashItemViewModel || BugBashItemHelpers.getNewBugBashItemViewModel(this.props.id)}
                    onItemAccept={(model: IBugBashItem, workItem: WorkItem) => {
                        let newViewModels = this.state.bugBashItemViewModels.slice();
                        let newViewModel = BugBashItemHelpers.getBugBashItemViewModel(model);
                        let newWorkItemsMap = {...this.state.workItemsMap};

                        let index = Utils_Array.findIndex(newViewModels, v => v.bugBashItem.id === model.id);
                        if (index !== -1) {
                            newViewModels[index] = newViewModel;
                        }
                        else {
                            newViewModels.push(newViewModel);
                        }

                        if (workItem) {
                            newWorkItemsMap[workItem.id] = workItem;
                            this.updateState({bugBashItemViewModels: newViewModels, selectedBugBashItemViewModel: null, workItemsMap: newWorkItemsMap});
                        }
                        else {
                            this.updateState({bugBashItemViewModels: newViewModels, selectedBugBashItemViewModel: newViewModel});
                        }
                    }}
                    onItemUpdate={(model: IBugBashItem) => {
                        if (model.id) {
                            let newViewModels = this.state.bugBashItemViewModels.slice();
                            let newVideModel = BugBashItemHelpers.getBugBashItemViewModel(model);

                            let index = Utils_Array.findIndex(newViewModels, v => v.bugBashItem.id === model.id);
                            if (index !== -1) {
                                newViewModels[index] = newVideModel;
                            }
                            else {
                                newViewModels.push(newVideModel);
                            }
                            this.updateState({bugBashItemViewModels: newViewModels, selectedBugBashItemViewModel: newVideModel});
                        }
                    }}
                    onDelete={(model: IBugBashItem) => {
                        if (model.id) {
                            let newViewModels = this.state.bugBashItemViewModels.slice();
                            let index = Utils_Array.findIndex(newViewModels, v => v.bugBashItem.id === model.id);
                            if (index !== -1) {
                                newViewModels.splice(index, 1);
                                this.updateState({bugBashItemViewModels: newViewModels, selectedBugBashItemViewModel: null});
                            }
                            else {
                                this.updateState({selectedBugBashItemViewModel: null});
                            }
                        }
                    }} 
                    onChange={(changedModel: IBugBashItem, newComment: string) => {
                        if (changedModel.id) {
                            let newViewModels = this.state.bugBashItemViewModels.slice();
                            let index = Utils_Array.findIndex(newViewModels, viewModel => viewModel.bugBashItem.id === changedModel.id);
                            if (index !== -1) {
                                newViewModels[index].bugBashItem = {...changedModel};
                                newViewModels[index].newComment = newComment;
                                this.updateState({bugBashItemViewModels: newViewModels, selectedBugBashItemViewModel: newViewModels[index]});
                            }
                        }
                    }} />
            </div>
        );
    }

    private _getExtraWorkItemGridColumns(): IExtraWorkItemGridColumn[] {
        let workItemIdToItemMap: IDictionaryNumberTo<IBugBashItem> = {};
        for (const viewModel of this.state.bugBashItemViewModels) {
            if (BugBashItemHelpers.isAccepted(viewModel.bugBashItem)) {
                workItemIdToItemMap[viewModel.bugBashItem.workItemId] = viewModel.bugBashItem;
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
                this.updateState({selectedBugBashItemViewModel: null});
            }
            else {
                this.updateState({selectedBugBashItemViewModel: viewModels[0]});
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
            if (!BugBashItemHelpers.isValid(viewModel.bugBashItem)) {
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
                            content={viewModel.bugBashItem.title}
                            delay={TooltipDelay.medium}
                            overflowMode={TooltipOverflowMode.Parent}
                            directionalHint={DirectionalHint.bottomLeftEdge}>

                            <Label className={`${getCellClassName(viewModel)}`}>
                                {`${BugBashItemHelpers.isDirty(viewModel) ? "* " : ""}${viewModel.bugBashItem.title}`}
                            </Label>
                        </TooltipHost>
                    )
                },
                sortFunction: (viewModel1: IBugBashItemViewModel, viewModel2: IBugBashItemViewModel, sortOrder: SortOrder) => {
                    let compareValue = Utils_String.ignoreCaseComparer(viewModel1.bugBashItem.title, viewModel2.bugBashItem.title);
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                },
                filterFunction: (viewModel: IBugBashItemViewModel, filterText: string) => Utils_String.caseInsensitiveContains(viewModel.bugBashItem.title, filterText)
            },
            {
                key: "team",
                name: "Team",
                minWidth: isRejectedGrid ? 100 : 200,
                maxWidth: isRejectedGrid ? 200 : 300,
                resizable: true,
                onRenderCell: (viewModel: IBugBashItemViewModel) => {
                    const team = StoresHub.teamStore.getItem(viewModel.bugBashItem.teamId);
                    return (
                        <TooltipHost 
                            content={team ? team.name : viewModel.bugBashItem.teamId}
                            delay={TooltipDelay.medium}
                            overflowMode={TooltipOverflowMode.Parent}
                            directionalHint={DirectionalHint.bottomLeftEdge}>

                            <Label className={`${getCellClassName(viewModel)}`}>
                                {team ? team.name : viewModel.bugBashItem.teamId}
                            </Label>
                        </TooltipHost>
                    )
                },
                sortFunction: (viewModel1: IBugBashItemViewModel, viewModel2: IBugBashItemViewModel, sortOrder: SortOrder) => {
                    const team1 = StoresHub.teamStore.getItem(viewModel1.bugBashItem.teamId);
                    const team2 = StoresHub.teamStore.getItem(viewModel2.bugBashItem.teamId);
                    const team1Name = team1 ? team1.name : viewModel1.bugBashItem.teamId;
                    const team2Name = team2 ? team2.name : viewModel2.bugBashItem.teamId;

                    let compareValue = Utils_String.ignoreCaseComparer(team1Name, team2Name);
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                },
                filterFunction: (viewModel: IBugBashItemViewModel, filterText: string) =>  {
                    const team = StoresHub.teamStore.getItem(viewModel.bugBashItem.teamId);
                    return Utils_String.caseInsensitiveContains(team ? team.name : viewModel.bugBashItem.teamId, filterText);
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
                            content={viewModel.bugBashItem.createdBy}
                            delay={TooltipDelay.medium}
                            directionalHint={DirectionalHint.bottomLeftEdge}>

                            <div className={getCellClassName(viewModel)}>
                                <IdentityView identityDistinctName={viewModel.bugBashItem.createdBy} />
                            </div>
                        </TooltipHost>
                    )
                },
                sortFunction: (viewModel1: IBugBashItemViewModel, viewModel2: IBugBashItemViewModel, sortOrder: SortOrder) => {
                    let compareValue = Utils_String.ignoreCaseComparer(viewModel1.bugBashItem.createdBy, viewModel2.bugBashItem.createdBy);
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                },
                filterFunction: (viewModel: IBugBashItemViewModel, filterText: string) => Utils_String.caseInsensitiveContains(viewModel.bugBashItem.createdBy, filterText)
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
                            content={Utils_Date.format(viewModel.bugBashItem.createdDate, "M/d/yyyy h:mm tt")}
                            delay={TooltipDelay.medium}
                            overflowMode={TooltipOverflowMode.Parent}
                            directionalHint={DirectionalHint.bottomLeftEdge}>

                            <Label className={`${getCellClassName(viewModel)}`}>
                                {Utils_Date.friendly(viewModel.bugBashItem.createdDate)}
                            </Label>
                        </TooltipHost>
                    )
                },
                sortFunction: (viewModel1: IBugBashItemViewModel, viewModel2: IBugBashItemViewModel, sortOrder: SortOrder) => {
                    let compareValue = Utils_Date.defaultComparer(viewModel1.bugBashItem.createdDate, viewModel2.bugBashItem.createdDate);
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
                            content={viewModel.bugBashItem.rejectedBy}
                            delay={TooltipDelay.medium}
                            directionalHint={DirectionalHint.bottomLeftEdge}>

                            <div className={getCellClassName(viewModel)}>
                                <IdentityView identityDistinctName={viewModel.bugBashItem.rejectedBy} />
                            </div>
                        </TooltipHost>
                    )
                },
                sortFunction: (viewModel1: IBugBashItemViewModel, viewModel2: IBugBashItemViewModel, sortOrder: SortOrder) => {
                    let compareValue = Utils_String.ignoreCaseComparer(viewModel1.bugBashItem.rejectedBy, viewModel2.bugBashItem.rejectedBy);
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                },
                filterFunction: (viewModel: IBugBashItemViewModel, filterText: string) => Utils_String.caseInsensitiveContains(viewModel.bugBashItem.rejectedBy, filterText)
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
                            content={viewModel.bugBashItem.rejectReason}
                            delay={TooltipDelay.medium}
                            overflowMode={TooltipOverflowMode.Parent}
                            directionalHint={DirectionalHint.bottomLeftEdge}>

                            <Label className={`${getCellClassName(viewModel)}`}>
                                {viewModel.bugBashItem.rejectReason}
                            </Label>
                        </TooltipHost>
                    )
                },
                sortFunction: (viewModel1: IBugBashItemViewModel, viewModel2: IBugBashItemViewModel, sortOrder: SortOrder) => {
                    let compareValue = Utils_String.ignoreCaseComparer(viewModel1.bugBashItem.rejectReason, viewModel2.bugBashItem.rejectReason);
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                },
                filterFunction: (viewModel: IBugBashItemViewModel, filterText: string) => Utils_String.caseInsensitiveContains(viewModel.bugBashItem.rejectReason, filterText)
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
        for (let viewModel of viewModels || this.state.bugBashItemViewModels) {
            if (BugBashItemHelpers.isDirty(viewModel)) {
                return true;
            }
        }

        return false;
    }
}