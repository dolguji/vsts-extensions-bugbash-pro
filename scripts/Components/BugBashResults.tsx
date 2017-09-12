import '../../css/BugBashResults.scss';

import * as React from 'react';

import { WorkItem } from "TFS/WorkItemTracking/Contracts";

import { Label } from "OfficeFabric/Label";
import { autobind } from "OfficeFabric/Utilities";
import { TooltipHost, TooltipDelay, DirectionalHint, TooltipOverflowMode } from "OfficeFabric/Tooltip";
import { SelectionMode, ISelection, Selection } from "OfficeFabric/utilities/selection";
import { IContextualMenuItem } from "OfficeFabric/components/ContextualMenu/ContextualMenu.Props";
import { Checkbox } from "OfficeFabric/Checkbox";

import { DateUtils } from "MB/Utils/Date";
import { StringUtils } from "MB/Utils/String";
import { CoreUtils } from "MB/Utils/Core";
import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "MB/Components/BaseComponent";
import { BaseStore } from "MB/Flux/Stores/BaseStore";
import { TeamActions } from "MB/Flux/Actions/TeamActions";
import { Loading } from "MB/Components/Loading";
import { Grid, SortOrder, GridColumn } from "MB/Components/Grid";
import { WorkItemGrid, ColumnPosition, IExtraWorkItemGridColumn } from "MB/Components/WorkItemGrid";
import { IdentityView } from "MB/Components/IdentityView";
import { Hub } from "MB/Components/Hub";

import SplitterLayout from "rc-split-layout";

import { confirmAction } from "../Helpers";
import { BugBashItemEditor } from "./BugBashItemEditor";
import { StoresHub } from "../Stores/StoresHub";
import { BugBashItemActions } from "../Actions/BugBashItemActions";
import { BugBashFieldNames, BugBashItemFieldNames, ResultsView } from '../Constants';
import { BugBash } from "../ViewModels/BugBash";
import { BugBashItem } from "../ViewModels/BugBashItem";
import { BugBashClientActionsHub } from "../Actions/ActionsHub";

interface IBugBashResultsState extends IBaseComponentState {
    bugBashItems: BugBashItem[];
    pendingBugBashItems: BugBashItem[];
    rejectedBugBashItems: BugBashItem[];
    acceptedWorkItemIds: number[];
    selectedBugBashItem?: BugBashItem;
    gridKeyCounter: number;
}

interface IBugBashResultsProps extends IBaseComponentProps {
    bugBash: BugBash;
    filterText?: string;
    view?: string;
}

export class BugBashResults extends BaseComponent<IBugBashResultsProps, IBugBashResultsState> {
    private _itemInvokedDelayedFunction: CoreUtils.DelayedFunction;
    private _selection: ISelection;

    constructor(props: IBugBashResultsProps, context?: any) {
        super(props, context);
        this._selection = new Selection({
            getKey: (item: any) => item.id,
            onSelectionChanged: () => {
                this._onBugBashItemSelectionChanged(this._selection.getSelection() as BugBashItem[]);
            }
        });
    }   

    protected initializeState() {
        this.state = {
            bugBashItems: null,
            pendingBugBashItems: null,
            rejectedBugBashItems: null,
            acceptedWorkItemIds: null,
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

        let selectedBugBashItem = null;
        if (this.state.selectedBugBashItem && this.state.selectedBugBashItem.id) {
            selectedBugBashItem = StoresHub.bugBashItemStore.getBugBashItem(this.props.bugBash.id, this.state.selectedBugBashItem.id);
        }
        if (selectedBugBashItem == null) {
            selectedBugBashItem = StoresHub.bugBashItemStore.getNewBugBashItem()
        }

        return {
            bugBashItems: bugBashItems,
            pendingBugBashItems: (bugBashItems || []).filter(item => !item.isAccepted && !item.getFieldValue<boolean>(BugBashItemFieldNames.Rejected, true)),
            rejectedBugBashItems: (bugBashItems || []).filter(item => !item.isAccepted && item.getFieldValue<boolean>(BugBashItemFieldNames.Rejected, true)),
            acceptedWorkItemIds: (bugBashItems || []).filter(item => item.isAccepted).map(item => item.workItemId),
            selectedBugBashItem: selectedBugBashItem,
            loading: StoresHub.teamStore.isLoading() || StoresHub.bugBashItemStore.isLoading(this.props.bugBash.id)
        } as IBugBashResultsState;
    }

    public componentDidMount() {
        super.componentDidMount();
        
        BugBashClientActionsHub.SelectedBugBashItemChanged.addListener(this._setSelectedItem);

        TeamActions.initializeTeams();
        BugBashItemActions.initializeItems(this.props.bugBash.id);
    }

    public componentWillUnmount() {
        super.componentWillUnmount();

        BugBashClientActionsHub.SelectedBugBashItemChanged.removeListener(this._setSelectedItem);
    }

    public componentWillReceiveProps(nextProps: Readonly<IBugBashResultsProps>) {
        if (this.props.bugBash.id !== nextProps.bugBash.id) {
            if (StoresHub.bugBashItemStore.isLoaded(nextProps.bugBash.id)) {
                const bugBashItems = StoresHub.bugBashItemStore.getBugBashItems(nextProps.bugBash.id);
                this.updateState({
                    bugBashItems: bugBashItems,
                    pendingBugBashItems: (bugBashItems || []).filter(item => !item.isAccepted && !item.getFieldValue<boolean>(BugBashItemFieldNames.Rejected, true)),
                    rejectedBugBashItems: (bugBashItems || []).filter(item => !item.isAccepted && item.getFieldValue<boolean>(BugBashItemFieldNames.Rejected, true)),
                    acceptedWorkItemIds: (bugBashItems || []).filter(item => item.isAccepted).map(item => item.workItemId),
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
                    pendingBugBashItems: null,
                    rejectedBugBashItems: null,
                    acceptedWorkItemIds: null,
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

    private _renderGrids(): JSX.Element {
        let pivotContent: JSX.Element;
        
        switch (this.props.view) {            
            case ResultsView.AcceptedItemsOnly:
                pivotContent = <WorkItemGrid
                    filterText={this.props.filterText}
                    selectionPreservedOnEmptyClick={true}
                    setKey={`bugbash-work-item-grid-${this.state.gridKeyCounter}`}
                    className="bugbash-item-grid"
                    workItemIds={this.state.acceptedWorkItemIds}
                    fieldRefNames={["System.Id", "System.Title", "System.State", "System.AssignedTo", "System.AreaPath"]}
                    noResultsText="No Accepted items"
                    extraColumns={this._getExtraWorkItemGridColumns()}
                />;                
                break;
            case ResultsView.RejectedItemsOnly:
                pivotContent = <Grid
                    filterText={this.props.filterText}
                    selectionPreservedOnEmptyClick={true}
                    setKey={`bugbash-rejected-item-grid-${this.state.gridKeyCounter}`}
                    className="bugbash-item-grid"
                    noResultsText="No Rejected items"
                    items={this.state.rejectedBugBashItems}
                    selectionMode={SelectionMode.none}
                    selection={this._selection}
                    getKey={(item: BugBashItem) => item.id}
                    columns={this._getBugBashItemGridColumns(true)}
                />;
                break;
            default:
                pivotContent = <Grid
                    filterText={this.props.filterText}
                    selectionPreservedOnEmptyClick={true}
                    setKey={`bugbash-pending-item-grid-${this.state.gridKeyCounter}`}
                    className="bugbash-item-grid"
                    items={this.state.pendingBugBashItems}
                    selectionMode={SelectionMode.none}
                    columns={this._getBugBashItemGridColumns(false)}
                    selection={this._selection}
                    getKey={(item: BugBashItem) => item.id}
                    noResultsText="No Pending items"
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
                        return <BugBashItemEditor bugBashId={this.props.bugBash.id} bugBashItem={this.state.selectedBugBashItem} />;
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

        let menuItems: IContextualMenuItem[] = [
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

        if (!this.state.selectedBugBashItem.isNew()) {
            menuItems.push({
                key: "delete", 
                iconProps: {iconName: "Cancel", style: { color: "#da0a00", fontWeight: "bold" }},
                title: "Delete item",
                onClick: this._deleteBugBashItem
            });
        }
        return menuItems;
    }

    @autobind
    private _acceptBugBashItem() {
        this.state.selectedBugBashItem.accept();
    }

    @autobind
    private _rejectBugBashItem() {
        this.state.selectedBugBashItem.setFieldValue(BugBashItemFieldNames.Rejected, !this.state.selectedBugBashItem.getFieldValue<boolean>(BugBashItemFieldNames.Rejected), false);
        this.state.selectedBugBashItem.setFieldValue(BugBashItemFieldNames.RejectedBy, `${VSS.getWebContext().user.name} <${VSS.getWebContext().user.uniqueName}>`, false);
        this.state.selectedBugBashItem.setFieldValue(BugBashItemFieldNames.RejectReason, "");        
    }

    @autobind
    private async _refreshBugBashItem() {
        if (!this.state.selectedBugBashItem.isNew()){
            const confirm = await confirmAction(this.state.selectedBugBashItem.isDirty(), "Refreshing the item will undo your unsaved changes. Are you sure you want to do that?");
            if (confirm) {
                this.state.selectedBugBashItem.refresh();
            }
        }
    }

    @autobind
    private async _revertBugBashItem() {
        const confirm = await confirmAction(true, "Are you sure you want to undo your changes to this item?");
        if (confirm) {
            this.state.selectedBugBashItem.reset();
        }
    }

    @autobind
    private async _deleteBugBashItem() {
        const confirm = await confirmAction(true, "Are you sure you want to delete this item?");
        if (confirm) {
            this.state.selectedBugBashItem.delete();
        }
    }

    @autobind
    private _saveSelectedItem() {
        this.state.selectedBugBashItem.save(this.props.bugBash.id);
    }

    private _getExtraWorkItemGridColumns(): IExtraWorkItemGridColumn[] {
        let workItemIdToItemMap: IDictionaryNumberTo<BugBashItem> = {};
        for (const bugBashItem of this.state.bugBashItems) {
            if (bugBashItem.isAccepted) {
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
                    onRenderCell: (workItem: WorkItem) => {
                        const createdBy = workItemIdToItemMap[workItem.id].getFieldValue<string>(BugBashItemFieldNames.CreatedBy, true);
                        return (
                            <TooltipHost 
                                content={createdBy}
                                delay={TooltipDelay.medium}
                                directionalHint={DirectionalHint.bottomLeftEdge}>

                                <IdentityView identityDistinctName={createdBy} />
                            </TooltipHost>
                        );
                    },
                    sortFunction: (workItem1: WorkItem, workItem2: WorkItem, sortOrder: SortOrder) => {
                        const createdBy1 = workItemIdToItemMap[workItem1.id].getFieldValue<string>(BugBashItemFieldNames.CreatedBy, true);
                        const createdBy2 = workItemIdToItemMap[workItem2.id].getFieldValue<string>(BugBashItemFieldNames.CreatedBy, true);
                        let compareValue = StringUtils.ignoreCaseComparer(createdBy1, createdBy2);
                        return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                    },
                    filterFunction: (workItem: WorkItem, filterText: string) => StringUtils.caseInsensitiveContains(workItemIdToItemMap[workItem.id].getFieldValue<string>(BugBashItemFieldNames.CreatedBy, true), filterText)
                }
            }
        ];
    }
    
    private _onBugBashItemSelectionChanged(bugBashItems: BugBashItem[]) {
        if (this._itemInvokedDelayedFunction) {
            this._itemInvokedDelayedFunction.cancel();
        }

        this._itemInvokedDelayedFunction = CoreUtils.delay(this, 100, () => {
            if (bugBashItems == null || bugBashItems.length !== 1) {
                this.updateState({selectedBugBashItem: StoresHub.bugBashItemStore.getNewBugBashItem()} as IBugBashResultsState);
            }
            else {
                this.updateState({selectedBugBashItem: bugBashItems[0]} as IBugBashResultsState);
            }
        });
    }

    @autobind
    private _setSelectedItem(bugBashItemId: string) {
        let selectedItem: BugBashItem = null;

        if (this._itemInvokedDelayedFunction) {
            this._itemInvokedDelayedFunction.cancel();
        }

        if (bugBashItemId) {
            selectedItem = StoresHub.bugBashItemStore.getBugBashItem(this.props.bugBash.id, bugBashItemId);
        }
    
        if (selectedItem) {
            if ((this._selection as any)._keyToIndexMap[selectedItem.id] >= 0) {
                this._selection.setKeySelected(selectedItem.id, true, true);
            }
            else {
                this.updateState({
                    selectedBugBashItem: selectedItem,
                    gridKeyCounter: this.state.gridKeyCounter + 1
                } as IBugBashResultsState);
            }
        }
        else {
            this.updateState({
                selectedBugBashItem: StoresHub.bugBashItemStore.getNewBugBashItem(),
                gridKeyCounter: this.state.gridKeyCounter + 1
            } as IBugBashResultsState);
        }
    }

    private _getBugBashItemGridColumns(isRejectedGrid: boolean): GridColumn[] {
        const gridCellClassName = "item-grid-cell";
        const getCellClassName = (bugBashItem: BugBashItem) => {
            let className = gridCellClassName;
            if (bugBashItem.isDirty()) {
                className += " is-dirty";
            }
            if (!bugBashItem.isValid()) {
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
                onRenderCell: (bugBashItem: BugBashItem) => {
                    const title = bugBashItem.getFieldValue<string>(BugBashItemFieldNames.Title);
                    return (
                        <TooltipHost 
                            content={title}
                            delay={TooltipDelay.medium}
                            overflowMode={TooltipOverflowMode.Parent}
                            directionalHint={DirectionalHint.bottomLeftEdge}>

                            <Label className={getCellClassName(bugBashItem)}>
                                {`${bugBashItem.isDirty() ? "* " : ""}${title}`}
                            </Label>
                        </TooltipHost>
                    )
                },
                filterFunction: (bugBashItem: BugBashItem, filterText: string) => StringUtils.caseInsensitiveContains(bugBashItem.getFieldValue<string>(BugBashItemFieldNames.Title), filterText),
                sortFunction: (bugBashItem1: BugBashItem, bugBashItem2: BugBashItem, sortOrder: SortOrder) => {
                    const title1 = bugBashItem1.getFieldValue<string>(BugBashItemFieldNames.Title);
                    const title2 = bugBashItem2.getFieldValue<string>(BugBashItemFieldNames.Title);
                    let compareValue = StringUtils.ignoreCaseComparer(title1, title2);
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                }
            },
            {
                key: "team",
                name: "Team",
                minWidth: isRejectedGrid ? 100 : 200,
                maxWidth: isRejectedGrid ? 200 : 300,
                resizable: true,
                onRenderCell: (bugBashItem: BugBashItem) => {
                    const teamId = bugBashItem.getFieldValue<string>(BugBashItemFieldNames.TeamId);
                    const team = StoresHub.teamStore.getItem(teamId);
                    return (
                        <TooltipHost 
                            content={team ? team.name : teamId}
                            delay={TooltipDelay.medium}
                            overflowMode={TooltipOverflowMode.Parent}
                            directionalHint={DirectionalHint.bottomLeftEdge}>

                            <Label className={getCellClassName(bugBashItem)}>
                                {team ? team.name : teamId}
                            </Label>
                        </TooltipHost>
                    )
                },
                sortFunction: (bugBashItem1: BugBashItem, bugBashItem2: BugBashItem, sortOrder: SortOrder) => {
                    const teamId1 = bugBashItem1.getFieldValue<string>(BugBashItemFieldNames.TeamId);
                    const teamId2 = bugBashItem2.getFieldValue<string>(BugBashItemFieldNames.TeamId);
                    const team1 = StoresHub.teamStore.getItem(teamId1);
                    const team2 = StoresHub.teamStore.getItem(teamId2);
                    const team1Name = team1 ? team1.name : teamId1;
                    const team2Name = team2 ? team2.name : teamId2;

                    let compareValue = StringUtils.ignoreCaseComparer(team1Name, team2Name);
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                },
                filterFunction: (bugBashItem: BugBashItem, filterText: string) =>  {
                    const teamId = bugBashItem.getFieldValue<string>(BugBashItemFieldNames.TeamId);
                    const team = StoresHub.teamStore.getItem(teamId);
                    return StringUtils.caseInsensitiveContains(team ? team.name : teamId, filterText);
                }
            },
            {
                key: "createdby",
                name: "Created By",
                minWidth: isRejectedGrid ? 100 : 200,
                maxWidth: isRejectedGrid ? 200 : 300,
                resizable: true,
                onRenderCell: (bugBashItem: BugBashItem) => {
                    const createdBy = bugBashItem.getFieldValue<string>(BugBashItemFieldNames.CreatedBy);
                    return (
                        <TooltipHost 
                            content={createdBy}
                            delay={TooltipDelay.medium}
                            directionalHint={DirectionalHint.bottomLeftEdge}>

                            <IdentityView className={getCellClassName(bugBashItem)} identityDistinctName={createdBy} />
                        </TooltipHost>
                    )
                },
                sortFunction: (bugBashItem1: BugBashItem, bugBashItem2: BugBashItem, sortOrder: SortOrder) => {
                    const createdBy1 = bugBashItem1.getFieldValue<string>(BugBashItemFieldNames.CreatedBy);
                    const createdBy2 = bugBashItem2.getFieldValue<string>(BugBashItemFieldNames.CreatedBy);
                    let compareValue = StringUtils.ignoreCaseComparer(createdBy1, createdBy2);
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                },
                filterFunction: (bugBashItem: BugBashItem, filterText: string) => StringUtils.caseInsensitiveContains(bugBashItem.getFieldValue<string>(BugBashItemFieldNames.CreatedBy), filterText)
            },
            {
                key: "createddate",
                name: "Created Date",
                minWidth: isRejectedGrid ? 100 : 200,
                maxWidth: isRejectedGrid ? 200 : 300,
                resizable: true,
                onRenderCell: (bugBashItem: BugBashItem) => {
                    const createdDate = bugBashItem.getFieldValue<Date>(BugBashItemFieldNames.CreatedDate);
                    return (
                        <TooltipHost 
                            content={DateUtils.format(createdDate, "m/d/yyyy h:mm tt")}
                            delay={TooltipDelay.medium}
                            directionalHint={DirectionalHint.bottomLeftEdge}>

                            <Label className={getCellClassName(bugBashItem)}>
                                {DateUtils.friendly(createdDate)}
                            </Label>
                        </TooltipHost>
                    )
                },
                sortFunction: (bugBashItem1: BugBashItem, bugBashItem2: BugBashItem, sortOrder: SortOrder) => {
                    const createdDate1 = bugBashItem1.getFieldValue<Date>(BugBashItemFieldNames.CreatedDate);
                    const createdDate2 = bugBashItem2.getFieldValue<Date>(BugBashItemFieldNames.CreatedDate);
                    let compareValue = DateUtils.defaultComparer(createdDate1, createdDate2);
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
                onRenderCell: (bugBashItem: BugBashItem) => {
                    const rejectedBy = bugBashItem.getFieldValue<string>(BugBashItemFieldNames.RejectedBy);
                    return (
                        <TooltipHost 
                            content={rejectedBy}
                            delay={TooltipDelay.medium}
                            directionalHint={DirectionalHint.bottomLeftEdge}>

                            <IdentityView className={getCellClassName(bugBashItem)} identityDistinctName={rejectedBy} />
                        </TooltipHost>
                    )
                },
                sortFunction: (bugBashItem1: BugBashItem, bugBashItem2: BugBashItem, sortOrder: SortOrder) => {
                    const rejectedBy1 = bugBashItem1.getFieldValue<string>(BugBashItemFieldNames.RejectedBy);
                    const rejectedBy2 = bugBashItem2.getFieldValue<string>(BugBashItemFieldNames.RejectedBy);
                    let compareValue = StringUtils.ignoreCaseComparer(rejectedBy1, rejectedBy2);
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                },
                filterFunction: (bugBashItem: BugBashItem, filterText: string) => StringUtils.caseInsensitiveContains(bugBashItem.getFieldValue<string>(BugBashItemFieldNames.RejectedBy), filterText)
            }, 
            {
                key: "rejectreason",
                name: "Reject Reason",
                minWidth: 200,
                maxWidth: 800,
                resizable: true,
                onRenderCell: (bugBashItem: BugBashItem) => {
                    const rejectReason = bugBashItem.getFieldValue<string>(BugBashItemFieldNames.RejectReason);
                    return (
                        <TooltipHost 
                            content={rejectReason}
                            delay={TooltipDelay.medium}
                            overflowMode={TooltipOverflowMode.Parent}
                            directionalHint={DirectionalHint.bottomLeftEdge}>

                            <Label className={getCellClassName(bugBashItem)}>
                                {rejectReason}
                            </Label>
                        </TooltipHost>
                    )
                },
                sortFunction: (bugBashItem1: BugBashItem, bugBashItem2: BugBashItem, sortOrder: SortOrder) => {
                    const rejectReason1 = bugBashItem1.getFieldValue<string>(BugBashItemFieldNames.RejectReason);
                    const rejectReason2 = bugBashItem2.getFieldValue<string>(BugBashItemFieldNames.RejectReason);
                    let compareValue = StringUtils.ignoreCaseComparer(rejectReason1, rejectReason2);
                    return sortOrder === SortOrder.DESC ? -1 * compareValue : compareValue;
                },
                filterFunction: (bugBashItem: BugBashItem, filterText: string) => StringUtils.caseInsensitiveContains(bugBashItem.getFieldValue<string>(BugBashItemFieldNames.RejectReason), filterText)
            });
        }

        return columns;
    }
}