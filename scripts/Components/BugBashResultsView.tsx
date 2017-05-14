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

import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { BaseStore } from "VSTS_Extension/Stores/BaseStore";
import { WorkItemFieldStore } from "VSTS_Extension/Stores/WorkItemFieldStore";
import { Loading } from "VSTS_Extension/Components/Common/Loading";
import { Grid } from "VSTS_Extension/Components/Grids/Grid";
import { WorkItemGrid } from "VSTS_Extension/Components/Grids/WorkItemGrid/WorkItemGrid";
import { openWorkItemDialog } from "VSTS_Extension/Components/Grids/WorkItemGrid/WorkItemGridHelpers";
import { IdentityView } from "VSTS_Extension/Components/WorkItemControls/IdentityView";
import { SortOrder, GridColumn } from "VSTS_Extension/Components/Grids/Grid.Props";

import { IBugBash, UrlActions, IBugBashItem, IBugBashItemModel } from "../Models";
import * as Helpers from "../Helpers";
import { BugBashItemManager } from "../BugBashItemManager";
import { BugBashItemEditor } from "./BugBashItemEditor";
import { BugBashStore } from "../Stores/BugBashStore";
import { StoresHub } from "../Stores/StoresHub";

enum FilterMode {
    All,
    AcceptedItems,
    PendingItems
}

function getTextFromFilterMode(filterMode: FilterMode) {
    switch (filterMode){
        case FilterMode.All:
            return "All";
        case FilterMode.AcceptedItems:
            return "Accepted items";
        default:
            return "Pending items";
    }
}

interface IBugBashResultsViewState extends IBaseComponentState {
    bugBashItem?: IBugBash;
    items?: IBugBashItemModel[];
    selectedItem?: IBugBashItemModel;
    filterMode?: FilterMode;
    hideFilterSelect?: boolean;
    workItemsMap?: IDictionaryNumberTo<WorkItem>;    
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
            items: null,
            selectedItem: null,
            filterMode: FilterMode.All,
            hideFilterSelect: false,
            workItemsMap: null
        };
    }

    protected async initialize() {        
        const found = await StoresHub.bugBashStore.ensureItem(this.props.id);

        if (!found) {
            this.updateState({
                bugBashItem: null,
                items: [],
                selectedItem: null
            });
        }
        else {
            StoresHub.workItemFieldStore.initialize();
            this._refreshData();
        }
    }   

    private async _refreshData() {
        this.updateState({workItemsMap: null, items: null, selectedItem: null});

        let items = await BugBashItemManager.beginGetItems(this.props.id);
        const workItemIds = items.filter(item => BugBashItemManager.isAccepted(item)).map(item => item.workItemId);
        const descriptionField = StoresHub.bugBashStore.getItem(this.props.id).itemDescriptionField;

        if (workItemIds.length > 0) {
            let workItems = await WitClient.getClient().getWorkItems(workItemIds, ["System.Id", "System.Title", "System.WorkItemType", "System.State", "System.AssignedTo", "System.AreaPath", descriptionField]);
            let map: IDictionaryNumberTo<WorkItem> = {};
            for (const workItem of workItems) {
                if (workItem) {
                    map[workItem.id] = workItem;
                }
            }
            this.updateState({workItemsMap: map, items: items.map(item => this._getItemModel(item))});
        }
        else {
            this.updateState({workItemsMap: {}, items: items.map(item => this._getItemModel(item))});
        }
    }

    private _getItemModel(model: IBugBashItem): IBugBashItemModel {
        return {
            model: BugBashItemManager.deepCopy(model),
            originalModel: BugBashItemManager.deepCopy(model),
            newComment: ""
        }
    }

    protected onStoreChanged() {
        const bugBashItem = StoresHub.bugBashStore.getItem(this.props.id);

        this.updateState({
            bugBashItem: bugBashItem,
            hideFilterSelect: (bugBashItem && bugBashItem.autoAccept) ? true : false,
            filterMode: (bugBashItem && bugBashItem.autoAccept) ? FilterMode.AcceptedItems : this.state.filterMode
        });
    }

    private _isDataLoading(): boolean {
        return !StoresHub.bugBashStore.isLoaded() || this.state.workItemsMap == null || !StoresHub.workItemFieldStore.isLoaded() || this.state.items == null;
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
                        
                        <div className="item-viewer">
                            <BugBashItemEditor 
                                itemModel={this.state.selectedItem || BugBashItemManager.getNewItemModel(this.props.id)}
                                onClickNew={() => {
                                    this.updateState({selectedItem: null});
                                }}
                                onItemUpdate={(item: IBugBashItem) => {
                                    if (item.id) {
                                        let newItems = this.state.items.slice();
                                        let newModel = BugBashItemManager.getItemModel(item);

                                        let index = Utils_Array.findIndex(newItems, i => i.model.id === item.id);
                                        if (index !== -1) {
                                            newItems[index] = newModel;
                                        }
                                        else {
                                            newItems.push(newModel);
                                        }
                                        this.updateState({items: newItems, selectedItem: newModel});
                                    }
                                }}
                                onDelete={(item: IBugBashItem) => {
                                    if (item.id) {
                                        let newItems = this.state.items.slice();
                                        let index = Utils_Array.findIndex(newItems, i => i.model.id === item.id);
                                        if (index !== -1) {
                                            newItems.splice(index, 1);
                                            this.updateState({items: newItems, selectedItem: null});
                                        }
                                        else {
                                            this.updateState({selectedItem: null});
                                        }
                                    }
                                }} 
                                onChange={(data: {id: string, title: string, description: string, newComment: string}) => {
                                    if (data.id) {
                                        let newItems = this.state.items.slice();
                                        let index = Utils_Array.findIndex(newItems, item => item.model.id === data.id);
                                        if (index !== -1) {
                                            newItems[index].model.title = data.title;
                                            newItems[index].model.description = data.description;
                                            newItems[index].newComment = data.newComment;
                                            this.updateState({items: newItems, selectedItem: newItems[index]});
                                        }
                                    }
                                }} />
                        </div>
                    </div>
                );
            }            
        }
    }

    private _renderGrid(): JSX.Element {
        const allItems = this.state.items.filter(item => !BugBashItemManager.isAccepted(item.model) || this.state.workItemsMap[item.model.workItemId] != null);

        if (this.state.filterMode === FilterMode.PendingItems) {
            const pendingItems = allItems.filter(item => !BugBashItemManager.isAccepted(item.model));
            return <Grid
                        className="bugbash-item-grid"
                        items={pendingItems}
                        columns={this._getGridColumns()}
                        commandBarProps={{menuItems: this._getCommandBarMenuItems(), farMenuItems: this._getCommandBarFarMenuItems()}}
                        contextMenuProps={{menuItems: this._getContextMenuItems}}
                        onItemInvoked={this._onItemInvoked}
                    />;
        }
        else if (this.state.filterMode === FilterMode.AcceptedItems) {
            const workItemIds = allItems.filter(item => BugBashItemManager.isAccepted(item.model)).map(item => item.model.workItemId);
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
            ]
            return <WorkItemGrid
                        className="bugbash-item-grid"
                        workItems={workItems}
                        fields={fields}
                        onWorkItemUpdated={(updatedWorkItem: WorkItem) => {
                            let map = {...this.state.workItemsMap};
                            map[updatedWorkItem.id] = updatedWorkItem;
                            this.updateState({workItemsMap: map});
                        }}
                        commandBarProps={{menuItems: this._getCommandBarMenuItems(), farMenuItems: this._getCommandBarFarMenuItems()}}
                        contextMenuProps={{menuItems: this._getContextMenuItems}}
                    />;
        }
        else {
            return <Grid
                        className="bugbash-item-grid"
                        items={allItems}
                        columns={this._getGridColumns()}
                        commandBarProps={{menuItems: this._getCommandBarMenuItems(), farMenuItems: this._getCommandBarFarMenuItems()}}
                        contextMenuProps={{menuItems: this._getContextMenuItems}}
                        onItemInvoked={this._onItemInvoked}
                    />;
        }
    }

    @autobind
    private async _onItemInvoked(item: IBugBashItemModel) {
        if (BugBashItemManager.isAccepted(item.model)) {
            this.updateState({selectedItem: null});
            let workItem = this.state.workItemsMap[item.model.workItemId];
            const updatedWorkItem: WorkItem = await openWorkItemDialog(null, workItem);
            if (updatedWorkItem && updatedWorkItem.rev > workItem.rev) {
                let map = {...this.state.workItemsMap};
                map[updatedWorkItem.id] = updatedWorkItem;
                this.updateState({workItemsMap: map});
            }
        }
        else {
            this.updateState({selectedItem: item});
        }
    }

    private _getItemTitle(item: IBugBashItemModel): string {
        let title = item.model.title;
        if (BugBashItemManager.isAccepted(item.model) && this.state.workItemsMap[item.model.workItemId] && this.state.workItemsMap[item.model.workItemId].fields["System.Title"]) {
            title = this.state.workItemsMap[item.model.workItemId].fields["System.Title"];
        }

        return title || "";
    }

    @autobind
    private _getGridColumns(): GridColumn[] {
        const gridCellClassName = "item-grid-cell";
        const getCellClassName = (item: IBugBashItemModel) => {
            let className = gridCellClassName;
            if (BugBashItemManager.isDirty(item)) {
                className += " is-dirty";
            }
            if (!BugBashItemManager.isValid(item.model)) {
                className += " is-invalid";
            }

            return className;
        }
        
        const idColumn: GridColumn = {
            key: "witid",
            name: "Work item id",
            minWidth: 80,
            maxWidth: 80,
            resizable: false,
            onRenderCell: (item: IBugBashItemModel) => {
                return <Label className={`${getCellClassName(item)} id-cell`}>
                            {BugBashItemManager.isAccepted(item.model) ? item.model.workItemId : (BugBashItemManager.isDirty(item) ? "*" : "")}
                        </Label>;
            },
            sortFunction: (item1: IBugBashItemModel, item2: IBugBashItemModel, sortOrder: SortOrder) => {
                const v1 = item1.model.workItemId;
                const v2 = item2.model.workItemId;
                let compareValue = (v1 > v2) ? 1 : -1;
                return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
            },
            filterFunction: (item: IBugBashItemModel, filterText: string) => {
                if (BugBashItemManager.isAccepted(item.model)) {
                    return Utils_String.caseInsensitiveContains(item.model.workItemId.toString(), filterText);
                }
                return false;
            }
        } as GridColumn;

        const acceptedColumns: GridColumn[] = [
            {
                key: "acceptedby",
                name: "Accepted By",
                minWidth: 100,
                maxWidth: 250,
                resizable: true,
                onRenderCell: (item: IBugBashItemModel) => {
                    return <div className={getCellClassName(item)}><IdentityView identityDistinctName={item.model.acceptedBy} /></div>;
                },
                sortFunction: (item1: IBugBashItemModel, item2: IBugBashItemModel, sortOrder: SortOrder) => {
                    let compareValue = Utils_String.ignoreCaseComparer(item1.model.acceptedBy, item2.model.acceptedBy);
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                },
                filterFunction: (item: IBugBashItemModel, filterText: string) => Utils_String.caseInsensitiveContains(item.model.acceptedBy, filterText)
            },
            {
                key: "accepteddate",
                name: "Accepted Date",
                minWidth: 80,
                maxWidth: 150,
                resizable: true,
                onRenderCell: (item: IBugBashItemModel) => {
                    const acceptedDate = item.model.acceptedDate;
                    return <Label className={getCellClassName(item)}>{acceptedDate ? Utils_Date.friendly(acceptedDate) : ""}</Label>;
                },
                sortFunction: (item1: IBugBashItemModel, item2: IBugBashItemModel, sortOrder: SortOrder) => {
                    const v1 = item1.model.acceptedDate;
                    const v2 = item2.model.acceptedDate;
                    let compareValue: number;
                    if (v1 == null && v2 == null) {
                        compareValue = 0;
                    }
                    else if (v1 == null && v2 != null) {
                        compareValue = -1;
                    }
                    else if (v1 != null && v2 == null) {
                        compareValue = 1;
                    }
                    else {
                        compareValue = Utils_Date.defaultComparer(v1, v2);
                    }
                    
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                }
            }
        ];

        let columns: GridColumn[] = [
            {
                key: "title",
                name: "Title",
                minWidth: 200,
                maxWidth: 800,
                resizable: true,
                onRenderCell: (item: IBugBashItemModel) => {                    
                    return <Label 
                                className={`${getCellClassName(item)} title-cell`} 
                                onClick={() => {
                                    this._onItemInvoked(item);
                                }
                            }>
                                {this._getItemTitle(item)}
                            </Label>
                },
                sortFunction: (item1: IBugBashItemModel, item2: IBugBashItemModel, sortOrder: SortOrder) => {
                    let compareValue = Utils_String.ignoreCaseComparer(this._getItemTitle(item1), this._getItemTitle(item2));
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                },
                filterFunction: (item: IBugBashItemModel, filterText: string) => Utils_String.caseInsensitiveContains(this._getItemTitle(item), filterText)
            },
            {
                key: "createdby",
                name: "Created By",
                minWidth: 100,
                maxWidth: 250,
                resizable: true,
                onRenderCell: (item: IBugBashItemModel) => {
                    return <div className={getCellClassName(item)}><IdentityView identityDistinctName={item.model.createdBy} /></div>;   
                },
                sortFunction: (item1: IBugBashItemModel, item2: IBugBashItemModel, sortOrder: SortOrder) => {
                    let compareValue = Utils_String.ignoreCaseComparer(item1.model.createdBy, item2.model.createdBy);
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                },
                filterFunction: (item: IBugBashItemModel, filterText: string) => Utils_String.caseInsensitiveContains(item.model.createdBy, filterText)
            },
            {
                key: "createddate",
                name: "Created Date",
                minWidth: 80,
                maxWidth: 150,
                resizable: true,
                onRenderCell: (item: IBugBashItemModel) => <Label className={getCellClassName(item)}>{Utils_Date.friendly(item.model.createdDate)}</Label>,
                sortFunction: (item1: IBugBashItemModel, item2: IBugBashItemModel, sortOrder: SortOrder) => {
                    let compareValue = Utils_Date.defaultComparer(item1.model.createdDate, item2.model.createdDate);
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                }
            }
        ];

        if (this.state.filterMode === FilterMode.All) {
            columns = [idColumn].concat(columns).concat(acceptedColumns);
        }
        
        return columns;
    }

    private _getCommandBarMenuItems(): IContextualMenuItem[] {
        return [
                {
                    key: "edit", name: "Edit", title: "Edit", iconProps: {iconName: "Edit"},
                    onClick: async (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                        const confirm = await Helpers.confirmAction(this._isAnyItemDirty(), "You have some unsaved items in the list. Navigating to a different page will remove all the unsaved data. Are you sure you want to do it?");
                        if (confirm) {
                            let navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
                            navigationService.updateHistoryEntry(UrlActions.ACTION_EDIT, {id: this.props.id});
                        }
                    }
                },
                {
                    key: "refresh", name: "Refresh", title: "Refresh list", iconProps: {iconName: "Refresh"},
                    onClick: async (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                        const confirm = await Helpers.confirmAction(this._isAnyItemDirty(), "You have some unsaved items in the list. Refreshing the page will remove all the unsaved data. Are you sure you want to do it?");
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
                    onClick: async (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                        const confirm = await Helpers.confirmAction(this._isAnyItemDirty(), "You have some unsaved items in the list. Navigating to a different page will remove all the unsaved data. Are you sure you want to do it?");
                        if (confirm) {
                            let navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
                            navigationService.updateHistoryEntry(UrlActions.ACTION_ALL, null);
                        }
                    }
                }
            ] as IContextualMenuItem[];

        const filterMenuItem = {
            key: "GridMode", name: `Showing: ${getTextFromFilterMode(this.state.filterMode)}`, title: "Select item filter", iconProps: {iconName: "Filter"},
            subMenuProps: {
                items: [
                    {key: "All", name: "Show all items"}, {key: "Pending", name: "Show only pending items"}, {key: "Accepted", name: "Show only accepted items"}
                ],
                onItemClick: (ev?: any, item?: IContextualMenuItem) => {
                    switch (item.key){
                        case "All":
                            this.updateState({filterMode: FilterMode.All});
                            break;
                        case "Pending":
                            this.updateState({filterMode: FilterMode.PendingItems});
                            break;
                        default:
                            this.updateState({filterMode: FilterMode.AcceptedItems});
                            break;
                    }
                }
            }
        } as IContextualMenuItem;

        if (!this.state.hideFilterSelect) {
            menuItems = [filterMenuItem].concat(menuItems);
        }

        return menuItems;
    }

    @autobind
    private _getContextMenuItems(selectedItems: IBugBashItemModel[]): IContextualMenuItem[] {
        return [            
            {
                key: "Delete", name: "Delete", title: "Delete selected items from the bug bash instance", iconProps: {iconName: "RemoveLink"}, 
                disabled: selectedItems.length === 0,
                onClick: async (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                    const confirm = await Helpers.confirmAction(true, "Are you sure you want to clear selected items from this bug bash? This action is irreversible. Any work item associated with a bug bash item will not be deleted.");
                    if (confirm) {
                        if (this.state.selectedItem && Utils_Array.findIndex(selectedItems, item => item.model.id === this.state.selectedItem.model.id) !== -1) {
                            this.updateState({selectedItem: null});
                        }
                        
                        await BugBashItemManager.deleteItems(selectedItems.map(i => i.model));
                        let newItems = this.state.items.slice();
                        newItems = Utils_Array.subtract(newItems, selectedItems, (i1, i2) => i1.model.id === i2.model.id ? 0 : 1);
                        
                        this.updateState({items: newItems});
                    }
                }
            }
        ];
    }

    private _isAnyItemDirty(): boolean {
        for (let item of this.state.items) {
            if (BugBashItemManager.isDirty(item)) {
                return true;
            }
        }

        return false;
    }
}