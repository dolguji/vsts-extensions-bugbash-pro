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

import { IBugBash, IBugBashItem, IBugBashItemViewModel } from "../Interfaces";
import { BugBashItemHelpers, confirmAction } from "../Helpers";
import { BugBashItemEditor } from "./BugBashItemEditor";
import { StoresHub } from "../Stores/StoresHub";
import { BugBashItemActions } from "../Actions/BugBashItemActions";

interface IBugBashResultsState extends IBaseComponentState {
    itemIdToIndexMap: IDictionaryStringTo<number>;
    bugBashItemViewModels: IBugBashItemViewModel[];
    selectedPivot: SelectedPivot;
    selectedBugBashItemViewModel?: IBugBashItemViewModel;
    loading?: boolean;
}

interface IBugBashResultsProps extends IBaseComponentProps {
    bugBash: IBugBash;
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
            itemIdToIndexMap: {},
            bugBashItemViewModels: null,
            selectedBugBashItemViewModel: BugBashItemHelpers.getNewViewModel(this.props.bugBash.id),
            selectedPivot: this.props.bugBash.autoAccept ? SelectedPivot.Accepted : SelectedPivot.Pending,
            loading: true
        };
    }

    protected getStores(): BaseStore<any, any, any>[] {
        return [StoresHub.bugBashItemStore, StoresHub.teamStore];
    }

    protected getStoresState(): IBugBashResultsState {
        const bugBashItems = StoresHub.bugBashItemStore.getBugBashItems(this.props.bugBash.id);

        return {
            bugBashItemViewModels: bugBashItems ? bugBashItems.map(item => BugBashItemHelpers.getViewModel(item)) : null,
            itemIdToIndexMap: this._prepareIdToIndexMap(bugBashItems),
            loading: StoresHub.teamStore.isLoading() || StoresHub.bugBashItemStore.isLoading(this.props.bugBash.id)
        } as IBugBashResultsState;
    }

    public componentDidMount() {
        super.componentDidMount();
        
        TeamActions.initializeTeams();
        BugBashItemActions.initializeItems(this.props.bugBash.id);
    }

    public componentWillReceiveProps(nextProps: Readonly<IBugBashResultsProps>): void {
        if (this.props.bugBash.id !== nextProps.bugBash.id) {
            if (StoresHub.bugBashItemStore.isLoaded(nextProps.bugBash.id)) {
                const bugBashItems = StoresHub.bugBashItemStore.getBugBashItems(nextProps.bugBash.id);
                this.updateState({
                    bugBashItemViewModels: bugBashItems.map(item => BugBashItemHelpers.getViewModel(item)),
                    itemIdToIndexMap: this._prepareIdToIndexMap(bugBashItems),
                    loading: false,
                    selectedBugBashItemViewModel: BugBashItemHelpers.getNewViewModel(nextProps.bugBash.id)
                } as IBugBashResultsState);
            }
            else {
                this.updateState({
                    loading: true,
                    bugBashItemViewModels: null,
                    itemIdToIndexMap: {},
                    selectedBugBashItemViewModel: BugBashItemHelpers.getNewViewModel(nextProps.bugBash.id)
                } as IBugBashResultsState);

                BugBashItemActions.initializeItems(nextProps.bugBash.id);
            }
        }        
    }

    private _prepareIdToIndexMap(bugBashItems: IBugBashItem[]): IDictionaryStringTo<number> {
        let map = {};
        if (bugBashItems == null) {
            return {};
        }
        else {
            for (let i = 0; i < bugBashItems.length; i++) {
                map[bugBashItems[i].id] = i;
            }
        }
        return map;
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
                { this.state.loading && <Loading /> }
                { this.state.bugBashItemViewModels && StoresHub.teamStore.isLoaded() &&
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
        const allBugBashItemViewModels = this.state.bugBashItemViewModels.filter(viewModel => !BugBashItemHelpers.isAccepted(viewModel.originalBugBashItem));
        const pendingBugBashItemViewModels = allBugBashItemViewModels.filter(viewModel => !BugBashItemHelpers.isAccepted(viewModel.originalBugBashItem) && !viewModel.originalBugBashItem.rejected);
        const rejectedBugBashItemViewModels = allBugBashItemViewModels.filter(viewModel => viewModel.originalBugBashItem.rejected);
        const acceptedWorkItemIds = allBugBashItemViewModels.filter(viewModel => BugBashItemHelpers.isAccepted(viewModel.originalBugBashItem)).map(viewModel => viewModel.originalBugBashItem.workItemId);

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
                            bugBashItem={this.state.selectedBugBashItemViewModel.updatedBugBashItem}
                            newComment={this.state.selectedBugBashItemViewModel.newComment}
                            save={() => {

                            }}
                            onChange={this._onItemChange} />;
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
            const isMenuDisabled = BugBashItemHelpers.isDirty(this.state.selectedBugBashItemViewModel) 
                || BugBashItemHelpers.isNew(this.state.selectedBugBashItemViewModel.originalBugBashItem);

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
                                    disabled={BugBashItemHelpers.isNew(this.state.selectedBugBashItemViewModel.originalBugBashItem)}
                                    className="reject-menu-item-checkbox"
                                    label="Reject"
                                    checked={this.state.selectedBugBashItemViewModel.updatedBugBashItem.rejected === true}
                                    onChange={(ev: React.FormEvent<HTMLElement>, isChecked: boolean) => {                                        
                                        let updatedItem = {...this.state.selectedBugBashItemViewModel.updatedBugBashItem};
                                        updatedItem.rejected = !updatedItem.rejected;
                                        updatedItem.rejectedBy = `${VSS.getWebContext().user.name} <${VSS.getWebContext().user.uniqueName}>`;
                                        updatedItem.rejectReason = "";
                                        this._onItemChange(updatedItem);
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

    @autobind
    private _onItemChange(changedItem: IBugBashItem, newComment?: string) {
        if (!BugBashItemHelpers.isNew(changedItem)) {
            const newViewModels = this.state.bugBashItemViewModels;
            const index = this.state.itemIdToIndexMap[changedItem.id];
            if (index >= 0 && index < this.state.bugBashItemViewModels.length) {
                newViewModels[index].updatedBugBashItem = {...changedItem};
                if (newComment != null) {
                    newViewModels[index].newComment = newComment;
                }
                this.updateState({
                    bugBashItemViewModels: newViewModels, 
                    selectedBugBashItemViewModel: newViewModels[index]
                } as IBugBashResultsState);
            }
        }
        else {
            let newSelectedBugBashItemViewModel = {...this.state.selectedBugBashItemViewModel};
            newSelectedBugBashItemViewModel.updatedBugBashItem = {...changedItem};
            if (newComment != null) {
                newSelectedBugBashItemViewModel.newComment = newComment;
            }

            this.updateState({
                selectedBugBashItemViewModel: newSelectedBugBashItemViewModel
            } as IBugBashResultsState);
        }
    }

    private _getItemEditorCommands(): IContextualMenuItem[] {
        return [
            {
                key: "save", name: "", 
                title: "Save", iconProps: {iconName: "Save"}, 
                disabled: !BugBashItemHelpers.isDirty(this.state.selectedBugBashItemViewModel)
                    || !BugBashItemHelpers.isValid(this.state.selectedBugBashItemViewModel),
                onClick: async () => {
                    //this._saveItem();
                }
            },
            {
                key: "refresh", name: "", 
                title: "Refresh", iconProps: {iconName: "Refresh"}, 
                disabled: BugBashItemHelpers.isNew(this.state.selectedBugBashItemViewModel.originalBugBashItem),
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
                disabled: !BugBashItemHelpers.isDirty(this.state.selectedBugBashItemViewModel),
                onClick: async () => {
                    const confirm = await confirmAction(true, "Are you sure you want to undo your changes to this item?");
                    if (confirm) {
                        this._onItemChange(this.state.selectedBugBashItemViewModel.originalBugBashItem, "");
                    }
                }
            }
        ];
    }

    private _getExtraWorkItemGridColumns(): IExtraWorkItemGridColumn[] {
        let workItemIdToItemMap: IDictionaryNumberTo<IBugBashItem> = {};
        for (const bugBashItemViewModel of this.state.bugBashItemViewModels) {
            if (BugBashItemHelpers.isAccepted(bugBashItemViewModel.originalBugBashItem)) {
                workItemIdToItemMap[bugBashItemViewModel.originalBugBashItem.workItemId] = bugBashItemViewModel.originalBugBashItem;
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
    private _onBugBashItemSelectionChanged(bugBashItemViewModels: IBugBashItemViewModel[]) {
        if (this._itemInvokedTimeout) {
            clearTimeout(this._itemInvokedTimeout);
            this._itemInvokedTimeout = null;
        }

        this._itemInvokedTimeout = setTimeout(() => {
            if (bugBashItemViewModels == null || bugBashItemViewModels.length !== 1) {
                this.updateState({selectedBugBashItemViewModel: BugBashItemHelpers.getNewViewModel(this.props.bugBash.id)} as IBugBashResultsState);
            }
            else {
                this.updateState({selectedBugBashItemViewModel: {...bugBashItemViewModels[0]}} as IBugBashResultsState);
            }

            this._itemInvokedTimeout = null;
        }, 200);
    }

    private _getBugBashItemGridColumns(isRejectedGrid: boolean): GridColumn[] {
        const gridCellClassName = "item-grid-cell";
        const getCellClassName = (bugBashItemViewModel: IBugBashItemViewModel) => {
            let className = gridCellClassName;
            if (BugBashItemHelpers.isDirty(bugBashItemViewModel)) {
                className += " is-dirty";
            }
            if (!BugBashItemHelpers.isValid(bugBashItemViewModel)) {
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
                onRenderCell: (bugBashItemViewModel: IBugBashItemViewModel) => {                 
                    return (
                        <TooltipHost 
                            content={bugBashItemViewModel.updatedBugBashItem.title}
                            delay={TooltipDelay.medium}
                            overflowMode={TooltipOverflowMode.Parent}
                            directionalHint={DirectionalHint.bottomLeftEdge}>

                            <Label className={`${getCellClassName(bugBashItemViewModel)}`}>
                                {`${BugBashItemHelpers.isDirty(bugBashItemViewModel) ? "* " : ""}${bugBashItemViewModel.updatedBugBashItem.title}`}
                            </Label>
                        </TooltipHost>
                    )
                },
                sortFunction: (viewModel1: IBugBashItemViewModel, viewModel2: IBugBashItemViewModel, sortOrder: SortOrder) => {
                    let compareValue = Utils_String.ignoreCaseComparer(viewModel1.updatedBugBashItem.title, viewModel2.updatedBugBashItem.title);
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                }
            },
            {
                key: "team",
                name: "Team",
                minWidth: isRejectedGrid ? 100 : 200,
                maxWidth: isRejectedGrid ? 200 : 300,
                resizable: true,
                onRenderCell: (bugBashItemViewModel: IBugBashItemViewModel) => {
                    const team = StoresHub.teamStore.getItem(bugBashItemViewModel.updatedBugBashItem.teamId);
                    return (
                        <TooltipHost 
                            content={team ? team.name : bugBashItemViewModel.updatedBugBashItem.teamId}
                            delay={TooltipDelay.medium}
                            overflowMode={TooltipOverflowMode.Parent}
                            directionalHint={DirectionalHint.bottomLeftEdge}>

                            <Label className={`${getCellClassName(bugBashItemViewModel)}`}>
                                {team ? team.name : bugBashItemViewModel.updatedBugBashItem.teamId}
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
                }
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
                }
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
                }
            });
        }

        return columns;
    }
}