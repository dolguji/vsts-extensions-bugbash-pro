import '../../css/BugBashResults.scss';

import * as React from 'react';

import Utils_Date = require('VSS/Utils/Date');
import Utils_String = require('VSS/Utils/String');
import { WorkItem } from "TFS/WorkItemTracking/Contracts";
import * as EventsService from 'VSS/Events/Services';
import { delay, DelayedFunction } from "VSS/Utils/Core";

import { Label } from "OfficeFabric/Label";
import { autobind } from "OfficeFabric/Utilities";
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

import { IBugBashItem } from "../Interfaces";
import { confirmAction } from "../Helpers";
import { BugBashItemEditor } from "./BugBashItemEditor";
import { StoresHub } from "../Stores/StoresHub";
import { BugBashItemActions } from "../Actions/BugBashItemActions";
import { BugBashFieldNames, BugBashItemFieldNames, Events, ResultsView } from '../Constants';
import { BugBash } from "../ViewModels/BugBash";
import { BugBashItem } from "../ViewModels/BugBashItem";

interface IBugBashResultsState extends IBaseComponentState {
    itemIdToIndexMap: IDictionaryStringTo<number>;
    bugBashItems: BugBashItem[];
    selectedBugBashItem?: BugBashItem;
    bugBashItemEditorError?: string;
    gridKeyCounter: number;
}

interface IBugBashResultsProps extends IBaseComponentProps {
    bugBash: BugBash;
    filterText?: string;
    view?: string;
}

export class BugBashResults extends BaseComponent<IBugBashResultsProps, IBugBashResultsState> {
    private _itemInvokedDelayedFunction: DelayedFunction;

    protected initializeState() {
        this.state = {
            itemIdToIndexMap: {},
            bugBashItems: null,
            selectedBugBashItem: StoresHub.bugBashItemStore.getNewBugBashItem(),
            loading: true,
            gridKeyCounter: 0
        };
    }

    protected getStores(): BaseStore<any, any, any>[] {
        return [StoresHub.bugBashItemStore, StoresHub.teamStore];
    }

    protected getStoresState(): IBugBashResultsState {
        const bugBashItems = StoresHub.bugBashItemStore.getBugBashItems(this.props.bugBash.id);

        return {
            bugBashItems: bugBashItems,
            itemIdToIndexMap: this._prepareIdToIndexMap(bugBashItems),
            loading: StoresHub.teamStore.isLoading() || StoresHub.bugBashItemStore.isLoading(this.props.bugBash.id)
        } as IBugBashResultsState;
    }

    public componentDidMount() {
        super.componentDidMount();
        
        EventsService.getService().attachEvent(Events.NewItem, this._clearSelectedItem);
        EventsService.getService().attachEvent(Events.RefreshItems, this._clearSelectedItem);
        TeamActions.initializeTeams();
        BugBashItemActions.initializeItems(this.props.bugBash.id);
    }

    public componentWillUnmount() {
        super.componentWillUnmount();

        EventsService.getService().detachEvent(Events.NewItem, this._clearSelectedItem);
        EventsService.getService().detachEvent(Events.RefreshItems, this._clearSelectedItem);
    }

    public componentWillReceiveProps(nextProps: Readonly<IBugBashResultsProps>) {
        if (this.props.bugBash.id !== nextProps.bugBash.id) {
            if (StoresHub.bugBashItemStore.isLoaded(nextProps.bugBash.id)) {
                const bugBashItems = StoresHub.bugBashItemStore.getBugBashItems(nextProps.bugBash.id);
                this.updateState({
                    bugBashItems: bugBashItems,
                    itemIdToIndexMap: this._prepareIdToIndexMap(bugBashItems),
                    loading: false,
                    bugBashItemEditorError: null,
                    selectedBugBashItem: StoresHub.bugBashItemStore.getNewBugBashItem(),
                    gridKeyCounter: this.state.gridKeyCounter + 1
                } as IBugBashResultsState);
            }
            else {
                this.updateState({
                    loading: true,
                    bugBashItems: null,
                    itemIdToIndexMap: {},
                    bugBashItemEditorError: null,
                    selectedBugBashItem: StoresHub.bugBashItemStore.getNewBugBashItem(),
                    gridKeyCounter: this.state.gridKeyCounter + 1
                } as IBugBashResultsState);

                BugBashItemActions.initializeItems(nextProps.bugBash.id);
            }
        }        
    }

    public render(): JSX.Element {
        if (this.state.loading) {
            return <Loading />;
        }
        return (
            <div className="bugbash-results">
                { this.state.bugBashItems && StoresHub.teamStore.isLoaded() &&
                    <SplitterLayout 
                        primaryIndex={0}
                        primaryMinSize={500}
                        secondaryMinSize={400}
                        secondaryInitialSize={500}
                        onChange={() => {
                            let evt = document.createEvent('UIEvents');
                            evt.initUIEvent('resize', true, false, window, 0);
                            window.dispatchEvent(evt);
                        }} >                        
                        { this._renderGrids() }
                        { this._renderItemEditor() }
                    </SplitterLayout>
                }
            </div>
        );
    }

    @autobind
    private _clearSelectedItem() {
        this.updateState({
            bugBashItemEditorError: null, 
            selectedBugBashItem: StoresHub.bugBashItemStore.getNewBugBashItem(),
            gridKeyCounter: this.state.gridKeyCounter + 1
        } as IBugBashResultsState);
    }

    private _prepareIdToIndexMap(bugBashItems: BugBashItem[]): IDictionaryStringTo<number> {
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

    private _renderGrids(): JSX.Element {
        let pivotContent: JSX.Element;
        
        switch (this.props.view) {            
            case ResultsView.AcceptedItemsOnly:
                const acceptedWorkItemIds = this.state.bugBashItems.filter(item => item.isAccepted).map(item => item.workItemId);
                pivotContent = <WorkItemGrid
                    filterText={this.props.filterText}
                    selectionPreservedOnEmptyClick={true}
                    setKey={`bugbash-work-item-grid-${this.state.gridKeyCounter}`}
                    className="bugbash-item-grid"
                    workItemIds={acceptedWorkItemIds}
                    fieldRefNames={["System.Id", "System.Title", "System.State", "System.AssignedTo", "System.AreaPath"]}
                    noResultsText="No Accepted items"
                    extraColumns={this._getExtraWorkItemGridColumns()}
                />;                
                break;
            case ResultsView.RejectedItemsOnly:
                const rejectedBugBashItemViewModels = this.state.bugBashItems.filter(item => !item.isAccepted && item.getFieldValue<boolean>(BugBashItemFieldNames.Rejected, true));
                pivotContent = <Grid
                    filterText={this.props.filterText}
                    selectionPreservedOnEmptyClick={true}
                    setKey={`bugbash-rejected-item-grid-${this.state.gridKeyCounter}`}
                    className="bugbash-item-grid"
                    noResultsText="No Rejected items"
                    items={rejectedBugBashItemViewModels}
                    selectionMode={SelectionMode.none}
                    columns={this._getBugBashItemGridColumns(true)}
                    events={{
                        onSelectionChanged: this._onBugBashItemSelectionChanged
                    }}
                />;
                break;
            default:
                const pendingBugBashItemViewModels = this.state.bugBashItems.filter(item => !item.isAccepted && !item.getFieldValue<boolean>(BugBashItemFieldNames.Rejected, true));
                pivotContent = <Grid
                    filterText={this.props.filterText}
                    selectionPreservedOnEmptyClick={true}
                    setKey={`bugbash-pending-item-grid-${this.state.gridKeyCounter}`}
                    className="bugbash-item-grid"
                    items={pendingBugBashItemViewModels}
                    selectionMode={SelectionMode.none}
                    columns={this._getBugBashItemGridColumns(false)}  
                    noResultsText="No Pending items"
                    events={{
                        onSelectionChanged: this._onBugBashItemSelectionChanged
                    }}
                />;
                break;
        }

        return (
            <div className="left-content">
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
                    onRenderPivotContent: () => {
                        return <BugBashItemEditor bugBashItem={this.state.selectedBugBashItem} />;
                    },
                    pivots: [
                        {
                            key: "addedititem",
                            text: this.state.selectedBugBashItem.isNew() ? "New Item" : "Edit Item",
                            commands: this._getItemEditorCommands(),
                            farCommands: this._getItemEditorFarCommands()
                        }
                    ]
                }}
            />
        </div>;
    }

    private _getItemEditorFarCommands(): IContextualMenuItem[] {
        let bugBash = StoresHub.bugBashStore.getItem(this.props.bugBash.id);

        if (!bugBash.getFieldValue<boolean>(BugBashFieldNames.AutoAccept, true)) {
            if (this.state.selectedBugBashItem.isAccepted) {
                return [];
            }

            const isMenuDisabled = this.state.selectedBugBashItem.isDirty() || this.state.selectedBugBashItem.isNew();

            return [
                {
                    key: "Accept", name: "Accept", title: "Create workitems from selected items", iconProps: {iconName: "Accept"}, className: !isMenuDisabled ? "acceptItemButton" : "",
                    disabled: isMenuDisabled,
                    onClick: this._acceptBugBashItem
                },
                {
                    key: "Reject",
                    onRender:() => {
                        return <Checkbox
                                    disabled={this.state.selectedBugBashItem.isNew()}
                                    className="reject-menu-item-checkbox"
                                    label="Reject"
                                    checked={this.state.selectedBugBashItem.getFieldValue<boolean>(BugBashItemFieldNames.Rejected)}
                                    onChange={this._rejectBugBashItem} />;
                    }
                }
            ];
        }
        else {
            return [{
                key: "Accept", name: "Auto accept on", className: "auto-accept-menuitem",
                title: "Auto accept is turned on for this bug bash. A work item would be created as soon as a bug bash item is created", 
                iconProps: {iconName: "SkypeCircleCheck"}
            }];
        }
    }    

    private _getItemEditorCommands(): IContextualMenuItem[] {
        if (this.state.selectedBugBashItem.isAccepted) {
            return [];
        }

        return [
            {
                key: "save", name: "", 
                iconProps: {iconName: "Save"}, 
                disabled: !this.state.selectedBugBashItem.isDirty()
                        || !this.state.selectedBugBashItem.isValid(),
                onClick: this._saveSelectedItem
            },
            {
                key: "refresh", name: "", 
                iconProps: {iconName: "Refresh"}, 
                disabled: this.state.selectedBugBashItem.isNew(),
                onClick: this._refreshBugBashItem
            },
            {
                key: "undo", name: "", 
                title: "Undo changes", iconProps: {iconName: "Undo"}, 
                disabled: !this.state.selectedBugBashItem.isDirty(),
                onClick: this._revertBugBashItem
            }
        ];
    }

    @autobind
    private _acceptBugBashItem() {
        this.state.selectedBugBashItem.accept();
    }

    @autobind
    private _rejectBugBashItem() {
        let updatedItem = {...this.state.selectedBugBashItem.updatedBugBashItem};
        updatedItem.rejected = !updatedItem.rejected;
        updatedItem.rejectedBy = `${VSS.getWebContext().user.name} <${VSS.getWebContext().user.uniqueName}>`;
        updatedItem.rejectReason = "";
        this._onItemChange(updatedItem);
    }

    @autobind
    private _refreshBugBashItem() {
        if (!BugBashItemHelpers.isNew(this.state.selectedBugBashItem.originalBugBashItem)){
            const confirm = await confirmAction(this.state.selectedBugBashItem.isDirty(), "Refreshing the item will undo your unsaved changes. Are you sure you want to do that?");
            if (confirm) {
                try {
                    const updatedItem = await BugBashItemActions.refreshItem(this.props.bugBash.id, this.state.selectedBugBashItem.originalBugBashItem.id);
        
                    this.updateState({bugBashItemEditorError: null, selectedBugBashItem: BugBashItemHelpers.getViewModel(updatedItem)} as IBugBashResultsState);
                }
                catch (e) {
                    this.updateState({bugBashItemEditorError: e} as IBugBashResultsState);
                }
            }
        }
    }

    @autobind
    private async _revertBugBashItem() {
        const confirm = await confirmAction(true, "Are you sure you want to undo your changes to this item?");
        if (confirm) {
            this._onItemChange(this.state.selectedBugBashItem.originalBugBashItem, "");
        }
    }

    @autobind
    private _onItemChange(changedItem: IBugBashItem, newComment?: string) {
        if (!BugBashItemHelpers.isNew(changedItem)) {
            const newViewModels = this.state.bugBashItems;
            const index = this.state.itemIdToIndexMap[changedItem.id];
            if (index >= 0 && index < this.state.bugBashItems.length) {
                newViewModels[index].updatedBugBashItem = {...changedItem};
                if (newComment != null) {
                    newViewModels[index].newComment = newComment;
                }
                this.updateState({
                    bugBashItems: newViewModels, 
                    selectedBugBashItem: newViewModels[index]
                } as IBugBashResultsState);
            }
        }
        else {
            let newSelectedBugBashItemViewModel = {...this.state.selectedBugBashItem};
            newSelectedBugBashItemViewModel.updatedBugBashItem = {...changedItem};
            if (newComment != null) {
                newSelectedBugBashItemViewModel.newComment = newComment;
            }

            this.updateState({
                selectedBugBashItem: newSelectedBugBashItemViewModel
            } as IBugBashResultsState);
        }
    }

    @autobind
    private async _saveSelectedItem() {
        if (this.state.selectedBugBashItem.isDirty() && BugBashItemHelpers.isValid(this.state.selectedBugBashItem)) {
            try {
                let updatedItem = await BugBashItemActions.saveItem(this.props.bugBash.id, this.state.selectedBugBashItem.updatedBugBashItem);
                
                if (this.props.bugBash.getFieldValue<boolean>(BugBashFieldNames.AutoAccept, true)) {
                    updatedItem = await BugBashItemActions.acceptBugBashItem(this.props.bugBash.id, updatedItem);
                }
                this.updateState({bugBashItemEditorError: null, selectedBugBashItem: BugBashItemHelpers.getViewModel(updatedItem)} as IBugBashResultsState);
            }
            catch (e) {
                this.updateState({bugBashItemEditorError: e} as IBugBashResultsState);
            }
        }
    }

    private _getExtraWorkItemGridColumns(): IExtraWorkItemGridColumn[] {
        let workItemIdToItemMap: IDictionaryNumberTo<IBugBashItem> = {};
        for (const bugBashItemViewModel of this.state.bugBashItems) {
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
                    onRenderCell: (workItem: WorkItem) => {
                        return (
                            <TooltipHost 
                                content={workItemIdToItemMap[workItem.id].createdBy}
                                delay={TooltipDelay.medium}
                                directionalHint={DirectionalHint.bottomLeftEdge}>

                                <IdentityView identityDistinctName={workItemIdToItemMap[workItem.id].createdBy} />
                            </TooltipHost>
                        );
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
    private _onBugBashItemSelectionChanged(bugBashItemViewModels: IBugBashItemViewModel[]) {
        if (this._itemInvokedDelayedFunction) {
            this._itemInvokedDelayedFunction.cancel();
        }

        this._itemInvokedDelayedFunction = delay(this, 100, () => {
            if (bugBashItemViewModels == null || bugBashItemViewModels.length !== 1) {
                this.updateState({bugBashItemEditorError: null, selectedBugBashItem: BugBashItemHelpers.getNewViewModel(this.props.bugBash.id)} as IBugBashResultsState);
            }
            else {
                this.updateState({bugBashItemEditorError: null, selectedBugBashItem: {...bugBashItemViewModels[0]}} as IBugBashResultsState);
            }
        });
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
                filterFunction: (viewModel: IBugBashItemViewModel, filterText: string) => Utils_String.caseInsensitiveContains(viewModel.updatedBugBashItem.title, filterText),
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

                            <IdentityView identityDistinctName={viewModel.updatedBugBashItem.createdBy} />
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

                            <IdentityView identityDistinctName={viewModel.updatedBugBashItem.rejectedBy} />
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
}