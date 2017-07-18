import "../../css/BugBashResults.scss";

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
import { Nav, INavLink, INavProps } from "OfficeFabric/Nav";
import { TooltipHost, TooltipDelay, DirectionalHint, TooltipOverflowMode } from "OfficeFabric/Tooltip";
import { SelectionMode } from "OfficeFabric/utilities/selection/interfaces";
import { CommandBar } from "OfficeFabric/CommandBar";
import { Overlay } from "OfficeFabric/Overlay";

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
    needsRefresh?: boolean;
}

interface IBugBashResultsProps extends IBaseComponentProps {
    bugBash: IBugBash;
}

enum SelectedPivot {
    Pending,
    Accepted,
    Rejected
}

export class BugBashResults extends BaseComponent<IBugBashResultsProps, IBugBashResultsState> {
    private _itemInvokedTimeout: any;

    protected initializeState() {
        this.state = {
            bugBashItemViewModels: null,
            selectedBugBashItemViewModel: null,
            workItemsMap: null,
            selectedPivot: this.props.bugBash.autoAccept ? SelectedPivot.Accepted : SelectedPivot.Pending,
            loading: true,
            needsRefresh: true
        };
    }

    protected getStores(): BaseStore<any, any, any>[] {
        return [StoresHub.bugBashItemStore, StoresHub.workItemFieldStore, StoresHub.teamStore];
    }

    protected getStoresState(): IBugBashResultsState {
        const bugBashItems = StoresHub.bugBashItemStore.getBugBashItems(this.props.bugBash.id);
        let workItemIds: number[];
        const isLoading = StoresHub.workItemFieldStore.isLoading() || StoresHub.teamStore.isLoading() || StoresHub.bugBashItemStore.isLoading(this.props.bugBash.id) || this.state.needsRefresh;

        let newState = {
            bugBashItemViewModels: bugBashItems ? bugBashItems.map(bugBashItem => BugBashItemHelpers.getBugBashItemViewModel(bugBashItem)) : null,
            loading: isLoading
        } as IBugBashResultsState

        if (this.state.needsRefresh && bugBashItems && !StoresHub.bugBashItemStore.isLoading(this.props.bugBash.id)) {
            workItemIds = bugBashItems.filter(bugBashItem => BugBashItemHelpers.isAccepted(bugBashItem)).map(bugBashItem => bugBashItem.workItemId);
            if (workItemIds.length > 0) {
                newState = {...newState, workItemsMap: null, selectedBugBashItemViewModel: null}
                this._refreshWorkItems(workItemIds);
            }
            else {
                newState = {...newState, workItemsMap: {}, needsRefresh: false, selectedBugBashItemViewModel: null, loading: false}
            }
        }

        return newState;
    }

    public componentDidMount() {
        super.componentDidMount();
        BugBashItemActions.initializeItems(this.props.bugBash.id);
        WorkItemFieldActions.initializeWorkItemFields();
        TeamActions.initializeTeams();
    }

    private async _refreshWorkItems(workItemIds: number[]) {
        if (workItemIds.length > 0) {
            let workItems = await WitClient.getClient().getWorkItems(workItemIds, ["System.Id", "System.Title", "System.WorkItemType", "System.State", "System.AssignedTo", "System.AreaPath"]);
            let map: IDictionaryNumberTo<WorkItem> = {};
            for (const workItem of workItems) {
                if (workItem) {
                    map[workItem.id] = workItem;
                }
            }
            this.updateState({workItemsMap: map, needsRefresh: false, loading: false} as IBugBashResultsState);
        }
    }

    private _getSelectedPivotKey(): string {
        switch (this.state.selectedPivot) {
            case SelectedPivot.Pending:
                return "Pending";
            case SelectedPivot.Accepted:
                return "Accepted";
            default:
                return "Rejected";
        }
    }

    private _updateSelectedPivot(pivotKey: string) {
        switch (pivotKey) {
            case "Pending":
                this.updateState({selectedPivot: SelectedPivot.Pending} as IBugBashResultsState);
                break;
            case "Accepted":
                this.updateState({selectedPivot: SelectedPivot.Accepted} as IBugBashResultsState);
                break;
           default:
                this.updateState({selectedPivot: SelectedPivot.Rejected} as IBugBashResultsState);
                break;
        }
    }
    
    public render(): JSX.Element {        
        return (
            <div className="bugbash-results">
                { this.state.loading && <Overlay><Loading /></Overlay> }
                { this.state.bugBashItemViewModels && this.state.workItemsMap && 
                    <SplitterLayout 
                        primaryIndex={0}
                        primaryMinSize={700}
                        secondaryMinSize={400}
                        secondaryInitialSize={500}
                        onChange={() => {
                            let evt = document.createEvent('UIEvents');
                            evt.initUIEvent('resize', true, false, window, 0);
                            window.dispatchEvent(evt);
                        }} >
                        
                        { this._renderPivots() }
                        { this._renderItemEditor() }
                    </SplitterLayout>
                }
            </div>
        );
    }

    private _renderPivots(): JSX.Element {
        const allBugBashItemViewModels = this.state.bugBashItemViewModels.filter(viewModel => !BugBashItemHelpers.isAccepted(viewModel.originalBugBashItem) || this.state.workItemsMap[viewModel.originalBugBashItem.workItemId] != null);
        const pendingBugBashItemViewModels = allBugBashItemViewModels.filter(viewModel => !BugBashItemHelpers.isAccepted(viewModel.originalBugBashItem) && !viewModel.originalBugBashItem.rejected);
        const rejectedBugBashItemViewModels = allBugBashItemViewModels.filter(viewModel => viewModel.originalBugBashItem.rejected);
        const acceptedWorkItemIds = allBugBashItemViewModels.filter(viewModel => BugBashItemHelpers.isAccepted(viewModel.originalBugBashItem)).map(viewModel => viewModel.originalBugBashItem.workItemId);

        let workItems: WorkItem[] = [];
        for (let id of acceptedWorkItemIds) {
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

        if (this.props.bugBash.autoAccept) {
            pivots = <Nav
                className="nav-panel"
                initialSelectedKey="Accepted"
                groups={ [{
                    links: [
                        { name: `Accepted Items (${workItems.length})`, key:"Accepted", url: "" },
                    ]
                }] }
            />;
        }
        else {
            pivots = <Nav
                className="nav-panel"
                initialSelectedKey={this._getSelectedPivotKey()}
                onLinkClick={(ev?: React.MouseEvent<HTMLElement>, item?: INavLink) => this._updateSelectedPivot(item.key)}
                groups={ [{
                    links: [
                        { name: `Pending Items (${pendingBugBashItemViewModels.length})`, key: "Pending", url: "" },
                        { name: `Accepted Items (${workItems.length})`, key:"Accepted", url: "" },
                        { name: `Rejected Items (${rejectedBugBashItemViewModels.length})`, key:"Rejected", url: "" }
                    ]
                }] }
            />;
        }
        
        switch (this.state.selectedPivot) {
            case SelectedPivot.Pending:
                pivotContent = <Grid
                    selectionPreservedOnEmptyClick={true}
                    setKey="bugbash-pending-item-grid"
                    className="bugbash-item-grid"
                    items={pendingBugBashItemViewModels}
                    selectionMode={SelectionMode.single}
                    columns={this._getBugBashItemGridColumns(false)}  
                    noResultsText="No Pending items"
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
                    commandBarProps={{
                        hideSearchBox: true,
                        hideCommandBar: true
                    }}
                    onWorkItemUpdated={(updatedWorkItem: WorkItem) => {
                        let map = {...this.state.workItemsMap};
                        map[updatedWorkItem.id] = updatedWorkItem;
                        this.updateState({workItemsMap: map} as IBugBashResultsState);
                    }}
                />;                
                break;
            default:
                pivotContent = <Grid
                    selectionPreservedOnEmptyClick={true}
                    setKey="bugbash-rejected-item-grid"
                    className="bugbash-item-grid"
                    noResultsText="No Rejected items"
                    items={rejectedBugBashItemViewModels}
                    selectionMode={SelectionMode.single}
                    columns={this._getBugBashItemGridColumns(true)}
                    events={{
                        onSelectionChanged: this._onBugBashItemSelectionChanged
                    }}
                />;
                break;            
        }

        return (
            <div className="left-content">
                {pivots}
                {pivotContent}
            </div>
        );
    }

    private _renderItemEditor(): JSX.Element {
        return (
            <div className="right-content">
                <div className="item-editor-container">
                    <div className="item-editor-container-header">
                        <Label className="item-editor-header overflow-ellipsis">{this.state.selectedBugBashItemViewModel ? "Edit Item" : "Add Item"}</Label>
                        <PrimaryButton    
                            className="add-new-item-button"                    
                            disabled={this.state.selectedBugBashItemViewModel == null}
                            text="New"
                            iconProps={ { iconName: "Add" } }
                            onClick={ () => this.updateState({selectedBugBashItemViewModel: null} as IBugBashResultsState) }
                            />
                    </div>
                    <BugBashItemEditor 
                        bugBashItemViewModel={this.state.selectedBugBashItemViewModel || BugBashItemHelpers.getNewBugBashItemViewModel(this.props.bugBash.id)}
                        onItemAccept={(bugBashItem: IBugBashItem, workItem: WorkItem) => {
                            let newViewModels = this.state.bugBashItemViewModels.slice();
                            let newViewModel = BugBashItemHelpers.getBugBashItemViewModel(bugBashItem);
                            let newWorkItemsMap = {...this.state.workItemsMap};

                            let index = Utils_Array.findIndex(newViewModels, v => v.updatedBugBashItem.id === bugBashItem.id);
                            if (index !== -1) {
                                newViewModels[index] = newViewModel;
                            }
                            else {
                                newViewModels.push(newViewModel);
                            }

                            if (workItem) {
                                newWorkItemsMap[workItem.id] = workItem;
                                this.updateState({bugBashItemViewModels: newViewModels, selectedBugBashItemViewModel: null, workItemsMap: newWorkItemsMap} as IBugBashResultsState);
                            }
                            else {
                                this.updateState({bugBashItemViewModels: newViewModels, selectedBugBashItemViewModel: newViewModel} as IBugBashResultsState);
                            }
                        }}
                        onItemUpdate={(bugBashItem: IBugBashItem) => {
                            if (bugBashItem.id) {
                                let newViewModels = this.state.bugBashItemViewModels.slice();
                                let newVideModel = BugBashItemHelpers.getBugBashItemViewModel(bugBashItem);

                                let index = Utils_Array.findIndex(newViewModels, v => v.updatedBugBashItem.id === bugBashItem.id);
                                if (index !== -1) {
                                    newViewModels[index] = newVideModel;
                                }
                                else {
                                    newViewModels.push(newVideModel);
                                }
                                this.updateState({bugBashItemViewModels: newViewModels, selectedBugBashItemViewModel: newVideModel} as IBugBashResultsState);
                            }
                        }}
                        onDelete={(bugBashItem: IBugBashItem) => {
                            if (bugBashItem.id) {
                                let newViewModels = this.state.bugBashItemViewModels.slice();
                                let index = Utils_Array.findIndex(newViewModels, v => v.updatedBugBashItem.id === bugBashItem.id);
                                if (index !== -1) {
                                    newViewModels.splice(index, 1);
                                    this.updateState({bugBashItemViewModels: newViewModels, selectedBugBashItemViewModel: null} as IBugBashResultsState);
                                }
                                else {
                                    this.updateState({selectedBugBashItemViewModel: null} as IBugBashResultsState);
                                }
                            }
                        }} 
                        onChange={(changedBugBashItem: IBugBashItem, newComment: string) => {
                            if (changedBugBashItem.id) {
                                let newViewModels = this.state.bugBashItemViewModels.slice();
                                let index = Utils_Array.findIndex(newViewModels, viewModel => viewModel.updatedBugBashItem.id === changedBugBashItem.id);
                                if (index !== -1) {
                                    newViewModels[index].updatedBugBashItem = {...changedBugBashItem};
                                    newViewModels[index].newComment = newComment;
                                    this.updateState({bugBashItemViewModels: newViewModels, selectedBugBashItemViewModel: newViewModels[index]} as IBugBashResultsState);
                                }
                            }
                        }} />
                </div>
            </div>
        );
    }

    private _getExtraWorkItemGridColumns(): IExtraWorkItemGridColumn[] {
        let workItemIdToItemMap: IDictionaryNumberTo<IBugBashItem> = {};
        for (const viewModel of this.state.bugBashItemViewModels) {
            if (BugBashItemHelpers.isAccepted(viewModel.updatedBugBashItem)) {
                workItemIdToItemMap[viewModel.updatedBugBashItem.workItemId] = viewModel.updatedBugBashItem;
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
                this.updateState({selectedBugBashItemViewModel: null} as IBugBashResultsState);
            }
            else {
                this.updateState({selectedBugBashItemViewModel: viewModels[0]} as IBugBashResultsState);
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
            if (!BugBashItemHelpers.isValid(viewModel.updatedBugBashItem)) {
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
                            content={viewModel.updatedBugBashItem.title}
                            delay={TooltipDelay.medium}
                            overflowMode={TooltipOverflowMode.Parent}
                            directionalHint={DirectionalHint.bottomLeftEdge}>

                            <Label className={`${getCellClassName(viewModel)}`}>
                                {`${BugBashItemHelpers.isDirty(viewModel) ? "* " : ""}${viewModel.updatedBugBashItem.title}`}
                            </Label>
                        </TooltipHost>
                    )
                },
                sortFunction: (viewModel1: IBugBashItemViewModel, viewModel2: IBugBashItemViewModel, sortOrder: SortOrder) => {
                    let compareValue = Utils_String.ignoreCaseComparer(viewModel1.updatedBugBashItem.title, viewModel2.updatedBugBashItem.title);
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                },
                filterFunction: (viewModel: IBugBashItemViewModel, filterText: string) => Utils_String.caseInsensitiveContains(viewModel.updatedBugBashItem.title, filterText)
            },
            {
                key: "team",
                name: "Team",
                minWidth: isRejectedGrid ? 100 : 200,
                maxWidth: isRejectedGrid ? 200 : 300,
                resizable: true,
                onRenderCell: (viewModel: IBugBashItemViewModel) => {
                    const team = StoresHub.teamStore.getItem(viewModel.updatedBugBashItem.teamId);
                    return (
                        <TooltipHost 
                            content={team ? team.name : viewModel.updatedBugBashItem.teamId}
                            delay={TooltipDelay.medium}
                            overflowMode={TooltipOverflowMode.Parent}
                            directionalHint={DirectionalHint.bottomLeftEdge}>

                            <Label className={`${getCellClassName(viewModel)}`}>
                                {team ? team.name : viewModel.updatedBugBashItem.teamId}
                            </Label>
                        </TooltipHost>
                    )
                },
                sortFunction: (viewModel1: IBugBashItemViewModel, viewModel2: IBugBashItemViewModel, sortOrder: SortOrder) => {
                    const team1 = StoresHub.teamStore.getItem(viewModel1.updatedBugBashItem.teamId);
                    const team2 = StoresHub.teamStore.getItem(viewModel2.updatedBugBashItem.teamId);
                    const team1Name = team1 ? team1.name : viewModel1.updatedBugBashItem.teamId;
                    const team2Name = team2 ? team2.name : viewModel2.updatedBugBashItem.teamId;

                    let compareValue = Utils_String.ignoreCaseComparer(team1Name, team2Name);
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                },
                filterFunction: (viewModel: IBugBashItemViewModel, filterText: string) =>  {
                    const team = StoresHub.teamStore.getItem(viewModel.updatedBugBashItem.teamId);
                    return Utils_String.caseInsensitiveContains(team ? team.name : viewModel.updatedBugBashItem.teamId, filterText);
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
                            content={viewModel.updatedBugBashItem.createdBy}
                            delay={TooltipDelay.medium}
                            directionalHint={DirectionalHint.bottomLeftEdge}>

                            <div className={getCellClassName(viewModel)}>
                                <IdentityView identityDistinctName={viewModel.updatedBugBashItem.createdBy} />
                            </div>
                        </TooltipHost>
                    )
                },
                sortFunction: (viewModel1: IBugBashItemViewModel, viewModel2: IBugBashItemViewModel, sortOrder: SortOrder) => {
                    let compareValue = Utils_String.ignoreCaseComparer(viewModel1.updatedBugBashItem.createdBy, viewModel2.updatedBugBashItem.createdBy);
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                },
                filterFunction: (viewModel: IBugBashItemViewModel, filterText: string) => Utils_String.caseInsensitiveContains(viewModel.updatedBugBashItem.createdBy, filterText)
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
                            content={Utils_Date.format(viewModel.updatedBugBashItem.createdDate, "M/d/yyyy h:mm tt")}
                            delay={TooltipDelay.medium}
                            overflowMode={TooltipOverflowMode.Parent}
                            directionalHint={DirectionalHint.bottomLeftEdge}>

                            <Label className={`${getCellClassName(viewModel)}`}>
                                {Utils_Date.friendly(viewModel.updatedBugBashItem.createdDate)}
                            </Label>
                        </TooltipHost>
                    )
                },
                sortFunction: (viewModel1: IBugBashItemViewModel, viewModel2: IBugBashItemViewModel, sortOrder: SortOrder) => {
                    let compareValue = Utils_Date.defaultComparer(viewModel1.updatedBugBashItem.createdDate, viewModel2.updatedBugBashItem.createdDate);
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
                            content={viewModel.updatedBugBashItem.rejectedBy}
                            delay={TooltipDelay.medium}
                            directionalHint={DirectionalHint.bottomLeftEdge}>

                            <div className={getCellClassName(viewModel)}>
                                <IdentityView identityDistinctName={viewModel.updatedBugBashItem.rejectedBy} />
                            </div>
                        </TooltipHost>
                    )
                },
                sortFunction: (viewModel1: IBugBashItemViewModel, viewModel2: IBugBashItemViewModel, sortOrder: SortOrder) => {
                    let compareValue = Utils_String.ignoreCaseComparer(viewModel1.updatedBugBashItem.rejectedBy, viewModel2.updatedBugBashItem.rejectedBy);
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                },
                filterFunction: (viewModel: IBugBashItemViewModel, filterText: string) => Utils_String.caseInsensitiveContains(viewModel.updatedBugBashItem.rejectedBy, filterText)
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
                            content={viewModel.updatedBugBashItem.rejectReason}
                            delay={TooltipDelay.medium}
                            overflowMode={TooltipOverflowMode.Parent}
                            directionalHint={DirectionalHint.bottomLeftEdge}>

                            <Label className={`${getCellClassName(viewModel)}`}>
                                {viewModel.updatedBugBashItem.rejectReason}
                            </Label>
                        </TooltipHost>
                    )
                },
                sortFunction: (viewModel1: IBugBashItemViewModel, viewModel2: IBugBashItemViewModel, sortOrder: SortOrder) => {
                    let compareValue = Utils_String.ignoreCaseComparer(viewModel1.updatedBugBashItem.rejectReason, viewModel2.updatedBugBashItem.rejectReason);
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                },
                filterFunction: (viewModel: IBugBashItemViewModel, filterText: string) => Utils_String.caseInsensitiveContains(viewModel.updatedBugBashItem.rejectReason, filterText)
            });
        }

        return columns;
    }

    private _isAnyBugBashItemViewModelDirty(viewModels?: IBugBashItemViewModel[]): boolean {
        for (let viewModel of viewModels || this.state.bugBashItemViewModels) {
            if (BugBashItemHelpers.isDirty(viewModel)) {
                return true;
            }
        }

        return false;
    }
}