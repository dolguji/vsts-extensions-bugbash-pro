import "../../css/BugBashResults.scss";

import * as React from "react";

import Utils_Date = require("VSS/Utils/Date");
import Utils_String = require("VSS/Utils/String");
import { WorkItem } from "TFS/WorkItemTracking/Contracts";

import { Label } from "OfficeFabric/Label";
import { autobind } from "OfficeFabric/Utilities";
import { Nav, INavLink } from "OfficeFabric/Nav";
import { TooltipHost, TooltipDelay, DirectionalHint, TooltipOverflowMode } from "OfficeFabric/Tooltip";
import { SelectionMode } from "OfficeFabric/utilities/selection/interfaces";
import { IContextualMenuItem } from "OfficeFabric/components/ContextualMenu/ContextualMenu.Props";
import { Checkbox } from "OfficeFabric/Checkbox";

import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { BaseStore } from "VSTS_Extension/Flux/Stores/BaseStore";
import { TeamActions } from "VSTS_Extension/Flux/Actions/TeamActions";
import { Loading } from "VSTS_Extension/Components/Common/Loading";
import { Grid } from "VSTS_Extension/Components/Grids/Grid";
import { WorkItemGrid } from "VSTS_Extension/Components/Grids/WorkItemGrid/WorkItemGrid";
import { IdentityView } from "VSTS_Extension/Components/WorkItemControls/IdentityView";
import { SortOrder, GridColumn } from "VSTS_Extension/Components/Grids/Grid.Props";
import { ColumnPosition, IExtraWorkItemGridColumn } from "VSTS_Extension/Components/Grids/WorkItemGrid/WorkItemGrid.Props";
import { Hub } from "VSTS_Extension/Components/Common/Hub/Hub";

import SplitterLayout from "rc-split-layout";

import { IBugBash, IBugBashItem } from "../Interfaces";
import { BugBashItemHelpers } from "../Helpers";
import { BugBashItemEditor } from "./BugBashItemEditor";
import { StoresHub } from "../Stores/StoresHub";
import { BugBashItemActions } from "../Actions/BugBashItemActions";
import { BugBashItemProvider } from "../BugBashItemProvider";

interface IBugBashResultsState extends IBaseComponentState {
    bugBashItems: IBugBashItem[];
    selectedPivot: SelectedPivot;
    selectedBugBashItem?: IBugBashItem;
    loading?: boolean;
    dirtyItemsMap?: IDictionaryStringTo<boolean>;
    invalidItemsMap?: IDictionaryStringTo<boolean>;
}

interface IBugBashResultsProps extends IBaseComponentProps {
    bugBash: IBugBash;
    provider: BugBashItemProvider;
    filterText?: string;
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
            bugBashItems: null,
            selectedBugBashItem: null,
            selectedPivot: this.props.bugBash.autoAccept ? SelectedPivot.Accepted : SelectedPivot.Pending,
            loading: true,
            dirtyItemsMap: {},
            invalidItemsMap: {}
        };
    }

    protected getStores(): BaseStore<any, any, any>[] {
        return [StoresHub.bugBashItemStore, StoresHub.teamStore];
    }

    protected getStoresState(): IBugBashResultsState {
        const bugBashItems = StoresHub.bugBashItemStore.getBugBashItems(this.props.bugBash.id);
        const isLoading = StoresHub.teamStore.isLoading() || StoresHub.bugBashItemStore.isLoading(this.props.bugBash.id);

        if (bugBashItems && !StoresHub.bugBashItemStore.isLoading(this.props.bugBash.id)) {
            this.props.provider.initialize(bugBashItems);
        }

        return {
            bugBashItems: bugBashItems,
            loading: isLoading
        } as IBugBashResultsState;
    }

    public componentDidMount() {
        super.componentDidMount();
        this.props.provider.addChangedListener(this._onProviderChange);
        
        TeamActions.initializeTeams();
        BugBashItemActions.initializeItems(this.props.bugBash.id);
    }

    public componentWillUnmount() {
        super.componentWillUnmount();
        this.props.provider.removeChangedListener(this._onProviderChange);
    }

    public componentWillReceiveProps(nextProps: Readonly<IBugBashResultsProps>): void {
        if (this.props.bugBash.id !== nextProps.bugBash.id) {
            BugBashItemActions.initializeItems(nextProps.bugBash.id);
        }        
    }

    @autobind
    private _onProviderChange() {
        let dirtyMap: IDictionaryStringTo<boolean> = {};
        let invalidMap: IDictionaryStringTo<boolean> = {};
        if (this.state.bugBashItems) {
            for (const bugBashItem of this.state.bugBashItems) {
                if (this.props.provider.isDirty(bugBashItem.id)) {
                    dirtyMap[bugBashItem.id] = true;
                }
                if (!this.props.provider.isValid(bugBashItem.id)) {
                    invalidMap[bugBashItem.id] = true;
                }
            }
        }

        this.updateState({dirtyItemsMap: dirtyMap, invalidItemsMap: invalidMap} as IBugBashResultsState);
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
                { this.state.bugBashItems &&
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
                        { this.state.loading && <Loading /> }
                        { this._renderPivots() }
                        { this._renderItemEditor() }
                    </SplitterLayout>
                }
            </div>
        );
    }

    private _renderPivots(): JSX.Element {
        const allBugBashItemViewModels = this.state.bugBashItems.filter(bugBashItem => !BugBashItemHelpers.isAccepted(bugBashItem));
        const pendingBugBashItemViewModels = allBugBashItemViewModels.filter(bugBashItem => !BugBashItemHelpers.isAccepted(bugBashItem) && !bugBashItem.rejected);
        const rejectedBugBashItemViewModels = allBugBashItemViewModels.filter(bugBash => bugBash.rejected);
        const acceptedWorkItemIds = allBugBashItemViewModels.filter(bugBash => BugBashItemHelpers.isAccepted(bugBash)).map(viewModel => viewModel.workItemId);

        let pivots: JSX.Element;
        let pivotContent: JSX.Element;

        if (this.props.bugBash.autoAccept) {
            pivots = <Nav
                className="nav-panel"
                initialSelectedKey="Accepted"
                groups={ [{
                    links: [
                        { name: `Accepted Items (${acceptedWorkItemIds.length})`, key:"Accepted", url: "" },
                    ]
                }] }
            />;
        }
        else {
            pivots = <Nav
                className="nav-panel"
                initialSelectedKey={this._getSelectedPivotKey()}
                onLinkClick={(_ev?: React.MouseEvent<HTMLElement>, item?: INavLink) => this._updateSelectedPivot(item.key)}
                groups={ [{
                    links: [
                        { name: `Pending Items (${pendingBugBashItemViewModels.length})`, key: "Pending", url: "" },
                        { name: `Accepted Items (${acceptedWorkItemIds.length})`, key:"Accepted", url: "" },
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
                    workItemIds={acceptedWorkItemIds}
                    fieldRefNames={["System.Id", "System.State", "System.AreaPath", "System.AssignedTo"]}
                    noResultsText="No Accepted items"
                    extraColumns={this._getExtraWorkItemGridColumns()}
                    commandBarProps={{
                        hideSearchBox: true,
                        hideCommandBar: true
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
        return <div className="right-content">
            <Hub 
                className="item-editor-container"
                title=""
                pivotProps={{
                    initialSelectedKey: "addedititem",
                    onRenderPivotContent: (key: string) => {
                        return <BugBashItemEditor 
                            bugBashItem={this.state.selectedBugBashItem || BugBashItemHelpers.getNewBugBashItem(this.props.bugBash.id)}
                            provider={this.props.provider} />;
                    },
                    pivots: [
                        {
                            key: "addedititem",
                            text: "Add item",
                            commands: this._getItemEditorCommands(),
                            farCommands: this._getItemEditorFarCommands()
                        }
                    ]
                }}
            />
        </div>;
    }

    private _getItemEditorFarCommands(): IContextualMenuItem[] {
        let menuItems: IContextualMenuItem[] = [];
        let bugBash = StoresHub.bugBashStore.getItem(this.props.bugBash.id);

        if (!bugBash.autoAccept) {
            const isMenuDisabled = false; //this.state.disableToolbar || BugBashItemHelpers.isDirty(this.state.bugBashItemViewModel) || this._isNew();
            menuItems.push(
                {
                    key: "Accept", name: "Accept", title: "Create workitems from selected items", iconProps: {iconName: "Accept"}, className: !isMenuDisabled ? "acceptItemButton" : "",
                    disabled: isMenuDisabled,
                    onClick: () => {
                        //this._acceptItem();
                    }
                },
                {
                    key: "Reject",
                    onRender:(item) => {
                        return <Checkbox
                                    disabled={false} //{this.state.disableToolbar || this._isNew()} 
                                    className="reject-menu-item-checkbox"
                                    label="Reject"
                                    checked={false} //{this.state.bugBashItemViewModel.updatedBugBashItem.rejected === true}
                                    onChange={(ev: React.FormEvent<HTMLElement>, isChecked: boolean) => {
                                        {/* let newViewModel = {...this.state.bugBashItemViewModel};
                                        newViewModel.updatedBugBashItem.rejected = !newViewModel.updatedBugBashItem.rejected;
                                        newViewModel.updatedBugBashItem.rejectedBy = `${VSS.getWebContext().user.name} <${VSS.getWebContext().user.uniqueName}>`;
                                        newViewModel.updatedBugBashItem.rejectReason = "";
                                        this.updateState({bugBashItemViewModel: newViewModel});
                                        this._onChange(newViewModel); */}
                                    }} />;
                    }
                });
        }
        else {
            menuItems.push({
                key: "Accept", name: "Auto accept on", className: "auto-accept-menuitem",
                title: "Auto accept is turned on for this bug bash. A work item would be created as soon as a bug bash item is created", 
                iconProps: {iconName: "SkypeCircleCheck"}
            });
        }

        return menuItems;
    }

    private _isItemDirty(bugBashItemId: string) {
        return this.state.dirtyItemsMap && this.state.dirtyItemsMap[bugBashItemId];
    }

    private _isItemInvalid(bugBashItemId: string) {
        return this.state.invalidItemsMap && this.state.invalidItemsMap[bugBashItemId];
    }

    private _getItemEditorCommands(): IContextualMenuItem[] {
        return [
            {
                key: "save", name: "", 
                title: "Save", iconProps: {iconName: "Save"}, 
                disabled: !this._isItemDirty(this.state.selectedBugBashItem ? this.state.selectedBugBashItem.id : "") 
                    || this._isItemInvalid(this.state.selectedBugBashItem ? this.state.selectedBugBashItem.id : ""), 
                onClick: async () => {
                    //this._saveItem();
                }
            },
            {
                key: "refresh", name: "", 
                title: "Refresh", iconProps: {iconName: "Refresh"}, 
                disabled: this.state.selectedBugBashItem == null,
                onClick: async () => {
                    // this.updateState({disableToolbar: true, error: null});
                    // let newModel: IBugBashItem;

                    // const confirm = await confirmAction(BugBashItemHelpers.isDirty(this.state.bugBashItemViewModel), "Refreshing the item will undo your unsaved changes. Are you sure you want to do that?");
                    // if (confirm) {
                    //     const id = this.state.bugBashItemViewModel.updatedBugBashItem.id;
                    //     const bugBashId = this.state.bugBashItemViewModel.updatedBugBashItem.bugBashId;

                    //     // this.updateState({viewModel: null});                        
                    //     // newModel = await BugBashItemManager.getItem(id, bugBashId);
                    //     // if (newModel) {
                    //     //     this.updateState({viewModel: BugBashItemHelpers.getItemViewModel(newModel), error: null, disableToolbar: false});
                    //     //     this.props.onItemUpdate(newModel);

                    //     //     StoresHub.bugBashItemCommentStore.refreshComments(id);
                    //     // }
                    //     // else {
                    //     //     this.updateState({viewModel: null, error: null, loadError: "This item no longer exist. Please refresh the list and try again."});
                    //     // }
                    // }
                    // else {
                    //     this.updateState({disableToolbar: false});                    
                    // }                
                }
            },
            {
                key: "undo", name: "", 
                title: "Undo changes", iconProps: {iconName: "Undo"}, 
                disabled: !this._isItemDirty(this.state.selectedBugBashItem ? this.state.selectedBugBashItem.id : ""),
                onClick: async () => {
                    // this.updateState({disableToolbar: true});

                    // const confirm = await confirmAction(true, "Are you sure you want to undo your changes to this item?");
                    // if (confirm) {
                    //     let newViewModel = {...this.state.bugBashItemViewModel};
                    //     newViewModel.updatedBugBashItem = {...newViewModel.originalBugBashItem};
                    //     newViewModel.newComment = "";
                    //     this.updateState({bugBashItemViewModel: newViewModel, disableToolbar: false, error: null});
                    //     this.props.onItemUpdate(newViewModel.updatedBugBashItem);
                    // }
                    // else {
                    //     this.updateState({disableToolbar: false});
                    // }
                }
            }
        ];
    }

    private _getExtraWorkItemGridColumns(): IExtraWorkItemGridColumn[] {
        let workItemIdToItemMap: IDictionaryNumberTo<IBugBashItem> = {};
        for (const bugBashItem of this.state.bugBashItems) {
            if (BugBashItemHelpers.isAccepted(bugBashItem)) {
                workItemIdToItemMap[bugBashItem.workItemId] = bugBashItem;
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
                    onRenderCell: (workItemId: number) => {
                        return <IdentityView identityDistinctName={workItemIdToItemMap[workItemId].createdBy} />;
                    },
                    sortFunction: (workItemId1: number, workItemId2: number, sortOrder: SortOrder) => {                        
                        let compareValue = Utils_String.ignoreCaseComparer(workItemIdToItemMap[workItemId1].createdBy, workItemIdToItemMap[workItemId2].createdBy);
                        return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                    },
                    filterFunction: (workItem: WorkItem, filterText: string) => Utils_String.caseInsensitiveContains(workItemIdToItemMap[workItem.id].createdBy, filterText)
                }
            }
        ];
    }
    
    @autobind
    private _onBugBashItemSelectionChanged(bugBashItems: IBugBashItem[]) {
        if (this._itemInvokedTimeout) {
            clearTimeout(this._itemInvokedTimeout);
            this._itemInvokedTimeout = null;
        }

        this._itemInvokedTimeout = setTimeout(() => {
            if (bugBashItems == null || bugBashItems.length !== 1) {
                this.updateState({selectedBugBashItem: null} as IBugBashResultsState);
            }
            else {
                this.updateState({selectedBugBashItem: bugBashItems[0]} as IBugBashResultsState);
            }

            this._itemInvokedTimeout = null;
        }, 200);
    }

    private _getBugBashItemGridColumns(isRejectedGrid: boolean): GridColumn[] {
        const gridCellClassName = "item-grid-cell";
        const getCellClassName = (bugBashItem: IBugBashItem) => {
            let className = gridCellClassName;
            if (this._isItemDirty(bugBashItem.id)) {
                className += " is-dirty";
            }
            if (this._isItemInvalid(bugBashItem.id)) {
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
                onRenderCell: (bugBashItem: IBugBashItem) => {                 
                    return (
                        <TooltipHost 
                            content={bugBashItem.title}
                            delay={TooltipDelay.medium}
                            overflowMode={TooltipOverflowMode.Parent}
                            directionalHint={DirectionalHint.bottomLeftEdge}>

                            <Label className={`${getCellClassName(bugBashItem)}`}>
                                {`${this._isItemDirty(bugBashItem.id) ? "* " : ""}${bugBashItem.title}`}
                            </Label>
                        </TooltipHost>
                    )
                },
                sortFunction: (item1: IBugBashItem, item2: IBugBashItem, sortOrder: SortOrder) => {
                    let compareValue = Utils_String.ignoreCaseComparer(item1.title, item2.title);
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                }
            },
            {
                key: "team",
                name: "Team",
                minWidth: isRejectedGrid ? 100 : 200,
                maxWidth: isRejectedGrid ? 200 : 300,
                resizable: true,
                onRenderCell: (item: IBugBashItem) => {
                    const team = StoresHub.teamStore.getItem(item.teamId);
                    return (
                        <TooltipHost 
                            content={team ? team.name : item.teamId}
                            delay={TooltipDelay.medium}
                            overflowMode={TooltipOverflowMode.Parent}
                            directionalHint={DirectionalHint.bottomLeftEdge}>

                            <Label className={`${getCellClassName(item)}`}>
                                {team ? team.name : item.teamId}
                            </Label>
                        </TooltipHost>
                    )
                },
                sortFunction: (item1: IBugBashItem, item2: IBugBashItem, sortOrder: SortOrder) => {
                    const team1 = StoresHub.teamStore.getItem(item1.teamId);
                    const team2 = StoresHub.teamStore.getItem(item2.teamId);
                    const team1Name = team1 ? team1.name : item1.teamId;
                    const team2Name = team2 ? team2.name : item2.teamId;

                    let compareValue = Utils_String.ignoreCaseComparer(team1Name, team2Name);
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                }
            },
            {
                key: "createdby",
                name: "Created By",
                minWidth: isRejectedGrid ? 100 : 200,
                maxWidth: isRejectedGrid ? 200 : 300,
                resizable: true,
                onRenderCell: (item: IBugBashItem) => {
                    return (
                        <TooltipHost 
                            content={item.createdBy}
                            delay={TooltipDelay.medium}
                            directionalHint={DirectionalHint.bottomLeftEdge}>

                            <div className={getCellClassName(item)}>
                                <IdentityView identityDistinctName={item.createdBy} />
                            </div>
                        </TooltipHost>
                    )
                },
                sortFunction: (item1: IBugBashItem, item2: IBugBashItem, sortOrder: SortOrder) => {
                    let compareValue = Utils_String.ignoreCaseComparer(item1.createdBy, item2.createdBy);
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                }
            },
            {
                key: "createddate",
                name: "Created Date",
                minWidth: isRejectedGrid ? 100 : 200,
                maxWidth: isRejectedGrid ? 200 : 300,
                resizable: true,
                onRenderCell: (item: IBugBashItem) => {
                    return (
                        <TooltipHost 
                            content={Utils_Date.format(item.createdDate, "M/d/yyyy h:mm tt")}
                            delay={TooltipDelay.medium}
                            overflowMode={TooltipOverflowMode.Parent}
                            directionalHint={DirectionalHint.bottomLeftEdge}>

                            <Label className={`${getCellClassName(item)}`}>
                                {Utils_Date.friendly(item.createdDate)}
                            </Label>
                        </TooltipHost>
                    )
                },
                sortFunction: (item1: IBugBashItem, item2: IBugBashItem, sortOrder: SortOrder) => {
                    let compareValue = Utils_Date.defaultComparer(item1.createdDate, item2.createdDate);
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
                onRenderCell: (item: IBugBashItem) => {
                    return (
                        <TooltipHost 
                            content={item.rejectedBy}
                            delay={TooltipDelay.medium}
                            directionalHint={DirectionalHint.bottomLeftEdge}>

                            <div className={getCellClassName(item)}>
                                <IdentityView identityDistinctName={item.rejectedBy} />
                            </div>
                        </TooltipHost>
                    )
                },
                sortFunction: (item1: IBugBashItem, item2: IBugBashItem, sortOrder: SortOrder) => {
                    let compareValue = Utils_String.ignoreCaseComparer(item1.rejectedBy, item2.rejectedBy);
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                }
            }, 
            {
                key: "rejectreason",
                name: "Reject Reason",
                minWidth: 200,
                maxWidth: 800,
                resizable: true,
                onRenderCell: (item: IBugBashItem) => {                 
                    return (
                        <TooltipHost 
                            content={item.rejectReason}
                            delay={TooltipDelay.medium}
                            overflowMode={TooltipOverflowMode.Parent}
                            directionalHint={DirectionalHint.bottomLeftEdge}>

                            <Label className={`${getCellClassName(item)}`}>
                                {item.rejectReason}
                            </Label>
                        </TooltipHost>
                    )
                },
                sortFunction: (item1: IBugBashItem, item2: IBugBashItem, sortOrder: SortOrder) => {
                    let compareValue = Utils_String.ignoreCaseComparer(item1.rejectReason, item2.rejectReason);
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                }
            });
        }

        return columns;
    }
}