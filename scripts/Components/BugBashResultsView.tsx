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

import { IBugBash, UrlActions, IBugBashItem } from "../Models";
import { BugBashItemEditor } from "./BugBashItemEditor";
import { BugBashStore } from "../Stores/BugBashStore";
import { BugBashItemStore } from "../Stores/BugBashItemStore";
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
    items?: IBugBashItem[];
    selectedBugBashItem?: IBugBashItem;
    loading?: boolean;
    filterMode?: FilterMode;
    hideFilterSelect?: boolean;
    workItemsMap?: IDictionaryNumberTo<WorkItem>;    
}

interface IBugBashResultsViewProps extends IBaseComponentProps {
    id: string;
}

export class BugBashResultsView extends BaseComponent<IBugBashResultsViewProps, IBugBashResultsViewState> {
    protected getStoresToLoad(): {new (): BaseStore<any, any, any>}[] {
        return [BugBashStore, BugBashItemStore, WorkItemFieldStore];
    }

    protected initializeState() {
        this.state = {
            bugBashItem: null,
            loading: true,
            items: [],
            selectedBugBashItem: null,
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
                loading: false,
                items: [],
                selectedBugBashItem: null
            });
        }
        else {
            this._refreshData();
        }
    }   

    private async _refreshData() {
        await StoresHub.bugBashItemStore.refreshItems(this.props.id);
        const items = StoresHub.bugBashItemStore.getItems(this.props.id);
        const workItemIds = items.filter(i => i.workItemId != null).map(i => i.workItemId);
        const descriptionField = StoresHub.bugBashStore.getItem(this.props.id).itemDescriptionField;

        if (workItemIds.length > 0) {
            let workItems = await WitClient.getClient().getWorkItems(workItemIds, ["System.Id", "System.Title", "System.WorkItemType", "System.State", "System.AssignedTo", "System.AreaPath", descriptionField]);
            let map: IDictionaryNumberTo<WorkItem> = {};
            for (const workItem of workItems) {
                map[workItem.id] = workItem;
            }
            this.updateState({workItemsMap: map});
        }
        else {
            this.updateState({workItemsMap: {}});
        }
    }

    protected onStoreChanged() {
        const bugBashItem = StoresHub.bugBashStore.getItem(this.props.id);

        this.updateState({
            bugBashItem: bugBashItem,
            hideFilterSelect: (bugBashItem && bugBashItem.autoAccept) ? true : false,
            filterMode: (bugBashItem && bugBashItem.autoAccept) ? FilterMode.AcceptedItems : this.state.filterMode,
            items: StoresHub.bugBashItemStore.getItems(this.props.id),
            loading: !StoresHub.bugBashStore.isLoaded() || this.state.workItemsMap == null || !StoresHub.workItemFieldStore.isLoaded() || !StoresHub.bugBashItemStore.isDataLoaded(this.props.id)
        });
    }

    public render(): JSX.Element {
        if (this.state.loading) {
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
                                id={this.state.selectedBugBashItem ? this.state.selectedBugBashItem.id : null} 
                                bugBashId={this.props.id} 
                                onClickNew={() => {
                                    this.updateState({selectedBugBashItem: null});
                                }} 
                                onSave={(item: IBugBashItem) => {
                                    this.updateState({selectedBugBashItem: item});
                                }}
                                onDelete={(item: IBugBashItem) => {
                                    this.updateState({selectedBugBashItem: null});                                    
                                }} />
                        </div>
                    </div>
                );
            }            
        }
    }

    private _renderGrid(): JSX.Element {
        const allItems = this.state.items.filter(i => i.workItemId == null || this.state.workItemsMap[i.workItemId] != null);

        if (this.state.filterMode === FilterMode.PendingItems) {
            const pendingItems = allItems.filter(i => i.workItemId == null);
            return <Grid
                        className="bugbash-item-grid"
                        items={pendingItems}
                        columns={this._getGridColumns()}
                        commandBarProps={{menuItems: this._getCommandBarMenuItems(), farMenuItems: this._getCommandBarFarMenuItems()}}
                        contextMenuProps={{menuItems: this._getContextMenuItems}}
                        onItemInvoked={(item: IBugBashItem) => this.updateState({selectedBugBashItem: item})}
                    />;
        }
        else if (this.state.filterMode === FilterMode.AcceptedItems) {
            const workItemIds = allItems.filter(i => i.workItemId != null).map(i => i.workItemId);
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
                        onItemInvoked={async (item: IBugBashItem) => {
                            if (item.workItemId) {
                                let workItem = this.state.workItemsMap[item.workItemId];
                                const updatedWorkItem: WorkItem = await openWorkItemDialog(null, this.state.workItemsMap[item.workItemId]);
                                if (updatedWorkItem && updatedWorkItem.rev > workItem.rev) {
                                    let map = {...this.state.workItemsMap};
                                    map[updatedWorkItem.id] = updatedWorkItem;
                                    this.updateState({workItemsMap: map});
                                }
                            }
                            else {
                                this.updateState({selectedBugBashItem: item});
                            }
                        }}
                    />;
        }
    }

    @autobind
    private _getGridColumns(): GridColumn[] {
        const gridCellClassName = "item-grid-cell";

        let columns: GridColumn[] = [            
            {
                key: "title",
                name: "Title",
                minWidth: 200,
                maxWidth: 800,
                resizable: true,
                onRenderCell: (item: IBugBashItem) => {
                    let text = item.title || "";
                    if (item.workItemId) {
                        text = this.state.workItemsMap[item.workItemId].fields["System.Title"] || item.title || "";
                    }

                    return <Label className={`${gridCellClassName} title-cell`}>
                                {text}
                            </Label>
                },
                sortFunction: (item1: IBugBashItem, item2: IBugBashItem, sortOrder: SortOrder) => {
                    let compareValue = Utils_String.ignoreCaseComparer(item1.title, item2.title)
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                },
                filterFunction: (item: IBugBashItem, filterText: string) => Utils_String.caseInsensitiveContains(item.title, filterText)
            },
            {
                key: "createdby",
                name: "Created By",
                minWidth: 100,
                maxWidth: 250,
                resizable: true,
                onRenderCell: (item: IBugBashItem) => <IdentityView identityDistinctName={item.createdBy} />,
                sortFunction: (item1: IBugBashItem, item2: IBugBashItem, sortOrder: SortOrder) => {
                    let compareValue = Utils_String.ignoreCaseComparer(item1.createdBy, item2.createdBy)
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                },
                filterFunction: (item: IBugBashItem, filterText: string) => Utils_String.caseInsensitiveContains(item.createdBy, filterText)
            },
            {
                key: "createddate",
                name: "Created Date",
                minWidth: 80,
                maxWidth: 150,
                resizable: true,
                onRenderCell: (item: IBugBashItem) => <Label className={gridCellClassName}>{Utils_Date.friendly(item.createdDate)}</Label>,
                sortFunction: (item1: IBugBashItem, item2: IBugBashItem, sortOrder: SortOrder) => {
                    let compareValue = Utils_Date.defaultComparer(item1.createdDate, item2.createdDate)
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                }
            },
            {
                key: "acceptedby",
                name: "Accepted By",
                minWidth: 100,
                maxWidth: 250,
                resizable: true,
                onRenderCell: (item: IBugBashItem) => <IdentityView identityDistinctName={item.acceptedBy} />,
                sortFunction: (item1: IBugBashItem, item2: IBugBashItem, sortOrder: SortOrder) => {
                    let compareValue = Utils_String.ignoreCaseComparer(item1.acceptedBy, item2.acceptedBy)
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                },
                filterFunction: (item: IBugBashItem, filterText: string) => Utils_String.caseInsensitiveContains(item.acceptedBy, filterText)
            },
            {
                key: "accepteddate",
                name: "Accepted Date",
                minWidth: 80,
                maxWidth: 150,
                resizable: true,
                onRenderCell: (item: IBugBashItem) => <Label className={gridCellClassName}>{Utils_Date.friendly(item.acceptedDate)}</Label>,
                sortFunction: (item1: IBugBashItem, item2: IBugBashItem, sortOrder: SortOrder) => {
                    let compareValue = Utils_Date.defaultComparer(item1.acceptedDate, item2.acceptedDate)
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                }
            }
        ];

        if (this.state.filterMode === FilterMode.All) {
            columns = [{
                key: "witid",
                name: "Work item id",
                minWidth: 80,
                maxWidth: 80,
                resizable: false,
                onRenderCell: (item: IBugBashItem) => {
                    return <Label className={`${gridCellClassName} id-cell`} onClick={() => this.updateState({selectedBugBashItem: item})}>
                                {item.workItemId || ""}
                            </Label>
                },
                sortFunction: (item1: IBugBashItem, item2: IBugBashItem, sortOrder: SortOrder) => {
                    let compareValue = Utils_String.ignoreCaseComparer(item1.workItemId.toString() || "", item2.workItemId.toString() || "")
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                },
                filterFunction: (item: IBugBashItem, filterText: string) => Utils_String.caseInsensitiveContains(item.workItemId.toString() , filterText)
            } as GridColumn].concat(columns);
        }

        return columns;
    }

    private _getCommandBarMenuItems(): IContextualMenuItem[] {
        return [
                {
                    key: "edit", name: "Edit", title: "Edit", iconProps: {iconName: "Edit"},
                    onClick: async (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                        let navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
                        navigationService.updateHistoryEntry(UrlActions.ACTION_EDIT, {id: this.props.id});
                    }
                },
                {
                    key: "refresh", name: "Refresh", title: "Refresh list", iconProps: {iconName: "Refresh"},
                    onClick: async (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {                        
                        await this._refreshData();
                        this.updateState({selectedBugBashItem: null});
                    }
                }
            ];
    }

    private _getCommandBarFarMenuItems(): IContextualMenuItem[] {
        let menuItems = [
                {
                    key: "Home", name: "Home", title: "Return to home view", iconProps: {iconName: "Home"}, 
                    onClick: async (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                        let navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
                        navigationService.updateHistoryEntry(UrlActions.ACTION_ALL, null);
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
    private _getContextMenuItems(selectedItems: IBugBashItem[]): IContextualMenuItem[] {
        return [            
            {
                key: "Delete", name: "Delete", title: "Delete selected items from the bug bash instance", iconProps: {iconName: "RemoveLink"}, 
                disabled: selectedItems.length === 0,
                onClick: async (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                    let dialogService: IHostDialogService = await VSS.getService(VSS.ServiceIds.Dialog) as IHostDialogService;
                    try {
                        await dialogService.openMessageDialog("Are you sure you want to clear selected items from this bug bash? This action is irreversible. Any work item associated with a bug bash item will not be deleted.", { useBowtieStyle: true });
                    }
                    catch (e) {
                        // user selected "No"" in dialog
                        return;
                    }

                    if (this.state.selectedBugBashItem && Utils_Array.findIndex(selectedItems, item => item.id === this.state.selectedBugBashItem.id) !== -1) {
                        this.updateState({selectedBugBashItem: null});
                    }
                    StoresHub.bugBashItemStore.deleteItems(selectedItems);
                }
            }
        ];
    }
}