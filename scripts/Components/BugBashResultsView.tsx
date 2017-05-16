import "../../css/BugBashResultsView.scss";

import * as React from "react";

import { HostNavigationService } from "VSS/SDK/Services/Navigation";
import Utils_Date = require("VSS/Utils/Date");
import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");
import { WorkItem, WorkItemField } from "TFS/WorkItemTracking/Contracts";
import * as WitClient from "TFS/WorkItemTracking/RestClient";

import { MessageBar, MessageBarType } from "OfficeFabric/MessageBar";
import { Label } from "OfficeFabric/Label";
import { IContextualMenuItem } from "OfficeFabric/components/ContextualMenu/ContextualMenu.Props";
import { autobind } from "OfficeFabric/Utilities";
import { Pivot, PivotItem } from "OfficeFabric/Pivot";

import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { BaseStore } from "VSTS_Extension/Stores/BaseStore";
import { WorkItemFieldStore } from "VSTS_Extension/Stores/WorkItemFieldStore";
import { Loading } from "VSTS_Extension/Components/Common/Loading";
import { Grid } from "VSTS_Extension/Components/Grids/Grid";
import { WorkItemGrid } from "VSTS_Extension/Components/Grids/WorkItemGrid/WorkItemGrid";
import { openWorkItemDialog } from "VSTS_Extension/Components/Grids/WorkItemGrid/WorkItemGridHelpers";
import { IdentityView } from "VSTS_Extension/Components/WorkItemControls/IdentityView";
import { SortOrder, GridColumn } from "VSTS_Extension/Components/Grids/Grid.Props";

import { UrlActions } from "../Constants";
import { IBugBash, IBugBashItem, IBugBashItemViewModel } from "../Interfaces";
import { confirmAction, BugBashItemHelpers } from "../Helpers";
import { BugBashItemManager } from "../BugBashItemManager";
import { BugBashItemEditor } from "./BugBashItemEditor";
import { BugBashStore } from "../Stores/BugBashStore";
import { StoresHub } from "../Stores/StoresHub";


interface IBugBashResultsViewState extends IBaseComponentState {
    bugBashItem?: IBugBash;
    viewModels?: IBugBashItemViewModel[];
    selectedViewModel?: IBugBashItemViewModel;
    workItemsMap?: IDictionaryNumberTo<WorkItem>;
    selectedPivot?: string;
}

interface IBugBashResultsViewProps extends IBaseComponentProps {
    id: string;
}

export class BugBashResultsView extends BaseComponent<IBugBashResultsViewProps, IBugBashResultsViewState> {
    protected getStoresToLoad(): {new (): BaseStore<any, any, any>}[] {
        return [BugBashStore, WorkItemFieldStore];
    }

    protected initializeState() {
        this.state = {
            bugBashItem: null,
            viewModels: null,
            selectedViewModel: null,
            workItemsMap: null
        };
    }

    protected async initialize() {        
        const found = await StoresHub.bugBashStore.ensureItem(this.props.id);

        if (!found) {
            this.updateState({
                bugBashItem: null,
                viewModels: [],
                selectedViewModel: null
            });
        }
        else {
            StoresHub.workItemFieldStore.initialize();
            this._refreshData();
        }
    }   

    private async _refreshData() {
        this.updateState({workItemsMap: null, viewModels: null, selectedViewModel: null});

        let items = await BugBashItemManager.beginGetItems(this.props.id);
        const workItemIds = items.filter(item => BugBashItemHelpers.isAccepted(item)).map(item => item.workItemId);
        const descriptionField = StoresHub.bugBashStore.getItem(this.props.id).itemDescriptionField;

        if (workItemIds.length > 0) {
            let workItems = await WitClient.getClient().getWorkItems(workItemIds, ["System.Id", "System.Title", "System.WorkItemType", "System.State", "System.AssignedTo", "System.AreaPath", descriptionField]);
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
            bugBashItem: bugBashItem
        });
    }

    private _isDataLoading(): boolean {
        return !StoresHub.bugBashStore.isLoaded() || this.state.workItemsMap == null || !StoresHub.workItemFieldStore.isLoaded() || this.state.viewModels == null;
    }
    
    public render(): JSX.Element {
        if (this._isDataLoading()) {
            return <Loading />;
        }
        else {
            if (!this.state.bugBashItem) {
                return <MessageBar messageBarType={MessageBarType.error}>This instance of bug bash doesn't exist.</MessageBar>;
            }
            else if(!Utils_String.equals(VSS.getWebContext().project.id, this.state.bugBashItem.projectId, true)) {
                return <MessageBar messageBarType={MessageBarType.error}>This instance of bug bash is out of scope of current project.</MessageBar>;
            }
            else {
                return (
                    <div className="results-view">
                        {this._renderGrid()}
                        {this._renderItemEditor()}
                    </div>
                );
            }            
        }
    }

    private _renderGrid(): JSX.Element {
        const allViewModels = this.state.viewModels.filter(viewModel => !BugBashItemHelpers.isAccepted(viewModel.model) || this.state.workItemsMap[viewModel.model.workItemId] != null);
        const pendingItems = allViewModels.filter(viewModel => !BugBashItemHelpers.isAccepted(viewModel.model));
        const workItemIds = allViewModels.filter(viewModel => BugBashItemHelpers.isAccepted(viewModel.model)).map(viewModel => viewModel.model.workItemId);
        let workItems: WorkItem[] = [];
        for (let id of workItemIds) {
            workItems.push(this.state.workItemsMap[id]);
        }

        const fields = [
            StoresHub.workItemFieldStore.getItem("System.Id"),
            StoresHub.workItemFieldStore.getItem("System.Title"),
            StoresHub.workItemFieldStore.getItem("System.WorkItemType"),
            StoresHub.workItemFieldStore.getItem("System.State"),
            StoresHub.workItemFieldStore.getItem("System.AssignedTo"),
            StoresHub.workItemFieldStore.getItem("System.AreaPath")
        ];

        const workItemGrid = <WorkItemGrid
                                className="bugbash-item-grid"
                                workItems={workItems}
                                fields={fields}
                                onWorkItemUpdated={(updatedWorkItem: WorkItem) => {
                                    let map = {...this.state.workItemsMap};
                                    map[updatedWorkItem.id] = updatedWorkItem;
                                    this.updateState({workItemsMap: map});
                                }}
                                commandBarProps={{menuItems: this._getCommandBarMenuItems(), farMenuItems: this._getCommandBarFarMenuItems()}}
                            />;
        
        if (!this.state.bugBashItem.autoAccept) {
            return (
                <div className="pivot-container">
                    <Pivot initialSelectedKey={this.state.selectedPivot || "Pending"} onLinkClick={(item: PivotItem) => this.updateState({selectedPivot: item.props.itemKey})}>
                        <PivotItem linkText="Pending Items" itemKey="Pending">
                            <Grid
                                className="bugbash-item-grid"
                                items={pendingItems}
                                columns={this._getPendingGridColumns()}
                                commandBarProps={{menuItems: this._getCommandBarMenuItems(), farMenuItems: this._getCommandBarFarMenuItems()}}
                                contextMenuProps={{menuItems: this._getContextMenuItems}}
                                onItemInvoked={this._onPendingItemInvoked}
                            />
                        </PivotItem>
                        <PivotItem linkText="AcceptedItems" itemKey="Accepted">
                            {workItemGrid}
                        </PivotItem>
                    </Pivot>
                </div>
            );
        }
        else {
            return workItemGrid;
        }        
    }

    private _renderItemEditor(): JSX.Element {
        return (
            <div className="item-editor-container">
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
                    onClickNew={() => {
                        this.updateState({selectedViewModel: null});
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
                    onChange={(changedData: {id: string, title: string, description: string}) => {
                        if (changedData.id) {
                            let newViewModels = this.state.viewModels.slice();
                            let index = Utils_Array.findIndex(newViewModels, viewModel => viewModel.model.id === changedData.id);
                            if (index !== -1) {
                                newViewModels[index].model.title = changedData.title;
                                newViewModels[index].model.description = changedData.description;
                                this.updateState({viewModels: newViewModels, selectedViewModel: newViewModels[index]});
                            }
                        }
                    }} />
            </div>
        );
    }

    @autobind
    private async _onPendingItemInvoked(viewModel: IBugBashItemViewModel) {
        this.updateState({selectedViewModel: viewModel});
    }

    @autobind
    private _getPendingGridColumns(): GridColumn[] {
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
                maxWidth: 800,
                resizable: true,
                onRenderCell: (viewModel: IBugBashItemViewModel) => {                    
                    return <Label 
                                className={`${getCellClassName(viewModel)} title-cell`} 
                                onClick={(e) => {
                                    this._onPendingItemInvoked(viewModel);
                                }
                            }>
                                {viewModel.model.title}
                            </Label>
                },
                sortFunction: (viewModel1: IBugBashItemViewModel, viewModel2: IBugBashItemViewModel, sortOrder: SortOrder) => {
                    let compareValue = Utils_String.ignoreCaseComparer(viewModel1.model.title, viewModel2.model.title);
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                },
                filterFunction: (viewModel: IBugBashItemViewModel, filterText: string) => Utils_String.caseInsensitiveContains(viewModel.model.title, filterText)
            },
            {
                key: "createdby",
                name: "Created By",
                minWidth: 100,
                maxWidth: 250,
                resizable: true,
                onRenderCell: (viewModel: IBugBashItemViewModel) => {
                    return <div className={getCellClassName(viewModel)}><IdentityView identityDistinctName={viewModel.model.createdBy} /></div>;   
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
                minWidth: 80,
                maxWidth: 150,
                resizable: true,
                onRenderCell: (viewModel: IBugBashItemViewModel) => <Label className={getCellClassName(viewModel)}>{Utils_Date.friendly(viewModel.model.createdDate)}</Label>,
                sortFunction: (viewModel1: IBugBashItemViewModel, viewModel2: IBugBashItemViewModel, sortOrder: SortOrder) => {
                    let compareValue = Utils_Date.defaultComparer(viewModel1.model.createdDate, viewModel2.model.createdDate);
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                }
            }
        ];

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

    @autobind
    private _getContextMenuItems(selectedViewModels: IBugBashItemViewModel[]): IContextualMenuItem[] {
        return [            
            {
                key: "Delete", name: "Delete", title: "Delete selected items from the bug bash instance", iconProps: {iconName: "RemoveLink"}, 
                disabled: selectedViewModels.length === 0,
                onClick: async () => {
                    const confirm = await confirmAction(true, "Are you sure you want to clear selected items from this bug bash? This action is irreversible. Any work item associated with a bug bash item will not be deleted.");
                    if (confirm) {
                        if (this.state.selectedViewModel && Utils_Array.findIndex(selectedViewModels, viewModel => viewModel.model.id === this.state.selectedViewModel.model.id) !== -1) {
                            this.updateState({selectedViewModel: null});
                        }
                        
                        await BugBashItemManager.deleteItems(selectedViewModels.map(v => v.model));
                        let newItems = this.state.viewModels.slice();
                        newItems = Utils_Array.subtract(newItems, selectedViewModels, (v1, v2) => v1.model.id === v2.model.id ? 0 : 1);
                        
                        this.updateState({viewModels: newItems});
                    }
                }
            }
        ];
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