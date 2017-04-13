import * as React from "react";

import { IBugBash, LoadingState, UrlActions, Constants } from "../Models";
import { HubView, IHubViewState } from "./HubView";
import { Loading } from "./Loading";
import { MessagePanel, MessageType } from "./MessagePanel";
import { WorkItemsViewer } from "./WorkItemsViewer";
import { NewWorkItemCreator } from "./NewWorkItemCreator";
import { WorkItemDiscussion } from "./WorkItemDiscussion";
import Helpers = require("../Helpers");

import { HostNavigationService } from "VSS/SDK/Services/Navigation";
import * as WitClient from "TFS/WorkItemTracking/RestClient";
import { Wiql, WorkItem } from "TFS/WorkItemTracking/Contracts";
import Utils_Date = require("VSS/Utils/Date");
import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");

import { CommandBar } from "OfficeFabric/CommandBar";
import { SearchBox } from "OfficeFabric/SearchBox";
import { IContextualMenuItem } from "OfficeFabric/components/ContextualMenu/ContextualMenu.Props";
import { autobind } from "OfficeFabric/Utilities";

interface IViewHubViewState extends IHubViewState {
    item: IBugBash;
    resultsLoaded: boolean;
    resultsLoading: boolean;
    workItemResults: WorkItem[];
    filterText: string;
    sortColumn: string;
    sortOrder: string;
    showDiscussions?: boolean;
    discussionWorkItem?: WorkItem;
}

export class ViewBugBashView extends HubView<IViewHubViewState> {

    public render(): JSX.Element {
        if (this.state.loadingState === LoadingState.Loading) {
            return <Loading />;
        }
        else {
            if (!this.state.item) {
                return <MessagePanel message="This instance of bug bash doesnt exist in the context of current team." messageType={MessageType.Error} />;
            }
            else {
                let rightSideComponent;

                if (this.state.showDiscussions && this.state.discussionWorkItem) {
                    rightSideComponent = <WorkItemDiscussion workItem={this.state.discussionWorkItem} onClose={this._hideDiscussions} context={this.props.context} />;
                }
                else if(this.state.item.templateId && !this.props.context.stores.workItemTemplateStore.getItem(this.state.item.templateId)) {
                    rightSideComponent = <MessagePanel message="The template specified in this instance of bug bash doesnt exist in the context of the current team." messageType={MessageType.Warning} />;
                }
                else {
                    rightSideComponent = <NewWorkItemCreator addWorkItem={this._addWorkItem} item={this.state.item} context={this.props.context} />;
                }

                return (
                    <div className="results-view">
                        <div className="results-view-menu">                            
                            <SearchBox className="results-view-searchbox" 
                                value={this.state.filterText || ""}
                                onSearch={(searchText: string) => this._updateFilterText(searchText)} 
                                onChange={(newText: string) => {
                                    if (newText.trim() === "") {
                                        this._updateFilterText("");
                                    }
                                }} />
                            <CommandBar className="results-view-menu-toolbar" items={this._getMenuItems()} 
                                farItems={
                                    [
                                        {
                                            key: "resultCount", name: `${this.state.workItemResults.length} workitems`, className: "result-count"
                                        },
                                        {
                                            key: "Home", name: "Home", title: "Return to home view", iconProps: {iconName: "Home"}, 
                                            onClick: async (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                                                let navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
                                                navigationService.updateHistoryEntry(UrlActions.ACTION_ALL, null);
                                            }
                                        }
                                    ]
                                } />
                        </div>
                        <div className="contents">
                            <WorkItemsViewer 
                                areResultsReady={!this.state.resultsLoading && this.state.resultsLoaded} 
                                sortColumn={this.state.sortColumn} 
                                sortOrder={this.state.sortOrder} 
                                bugBashItem={this.state.item}
                                workItems={this._sortAndFilterWorkItems(this.state.workItemResults)} 
                                refreshWorkItems={this._refreshWorkItemsView} 
                                changeSort={this._changeSort}
                                onShowDiscussions={this._showDiscussions}
                                updateWorkItems={this._addOrEditWorkItemsInView}
                                context={this.props.context} />
                            
                            <div className="right-side-content">
                                {rightSideComponent}
                            </div>
                        </div>
                    </div>
                );                
            }            
        }
    }

    protected async initialize() {                
        let found = await this.props.context.actionsCreator.ensureBugBash(this.props.id);
        if (!found) {
            this.setState({
                item: null,
                loadingState: LoadingState.Loaded,
                resultsLoading: false,
                resultsLoaded: false,
                workItemResults: [],
                filterText: ""
            });
        }
        else {
            this.props.context.actionsCreator.initializeWorkItemFields();
            this.props.context.actionsCreator.initializeWorkItemTemplates();
            this._refreshWorkItemResults();
        }
    }

    protected getStateFromStore(): IViewHubViewState {
        const item = this.props.context.stores.bugBashItemStore.getItem(this.props.id);
        let loadingState = this.props.context.stores.bugBashItemStore.isLoaded() && this.props.context.stores.workItemFieldStore.isLoaded() && this.props.context.stores.workItemTemplateStore.isLoaded() 
                ? LoadingState.Loaded : LoadingState.Loading;
        
        return this._mergeState({item: item, loadingState: loadingState});
    }

    private _getDefaultState(): IViewHubViewState {
        return {
            item: null,
            loadingState: LoadingState.Loading,
            resultsLoaded: false,
            resultsLoading: false,
            workItemResults: [],
            sortColumn: "System.CreatedDate",
            sortOrder: "desc",
            filterText: ""
        }
    }

    @autobind
    private _sortAndFilterWorkItems(workItems: WorkItem[]): WorkItem[] {
        let items = workItems.slice();
        let sortedItems = items.sort((w1: WorkItem, w2: WorkItem) => {
            if (Utils_String.equals(this.state.sortColumn, "ID", true)) {
                return this.state.sortOrder === "desc" ? ((w1.id > w2.id) ? -1 : 1) : ((w1.id > w2.id) ? 1 : -1);
            }
            else if (Utils_String.equals(this.state.sortColumn, "System.CreatedDate", true)) {
                let d1 = new Date(w1.fields["System.CreatedDate"]);
                let d2 = new Date(w2.fields["System.CreatedDate"]);
                return this.state.sortOrder === "desc" ? -1 * Utils_Date.defaultComparer(d1, d2) : Utils_Date.defaultComparer(d1, d2);
            }
            else if (Utils_String.equals(this.state.sortColumn, Constants.ACCEPT_STATUS_CELL_NAME, true)) {
                let v1 = Helpers.isWorkItemAccepted(w1) ? Constants.ACCEPTED_TEXT : (Helpers.isWorkItemRejected(w1) ? Constants.REJECTED_TEXT : "");
                let v2 = Helpers.isWorkItemAccepted(w2) ? Constants.ACCEPTED_TEXT : (Helpers.isWorkItemRejected(w2) ? Constants.REJECTED_TEXT : "");
                return this.state.sortOrder === "desc" ? -1 * Utils_String.ignoreCaseComparer(v1, v2) : Utils_String.ignoreCaseComparer(v1, v2);
            }
            else {
                let v1 = w1.fields[this.state.sortColumn];
                let v2 = w2.fields[this.state.sortColumn];
                return this.state.sortOrder === "desc" ? -1 * Utils_String.ignoreCaseComparer(v1, v2) : Utils_String.ignoreCaseComparer(v1, v2);
            }
        });

        if (!this.state.filterText) {
            return sortedItems;
        }
        else {
            return sortedItems.filter((workItem: WorkItem) => {
                let status = Helpers.isWorkItemAccepted(workItem) ? Constants.ACCEPTED_TEXT : (Helpers.isWorkItemRejected(workItem) ? Constants.REJECTED_TEXT : "");
                const filterText = this.state.filterText;
                return `${workItem.id}` === filterText
                    || Utils_String.caseInsensitiveContains(workItem.fields["System.AssignedTo"] || "", filterText)
                    || Utils_String.caseInsensitiveContains(workItem.fields["System.State"] || "", filterText)
                    || Utils_String.caseInsensitiveContains(workItem.fields["System.CreatedBy"] || "", filterText)
                    || Utils_String.caseInsensitiveContains(workItem.fields["System.Title"] || "", filterText)
                    || Utils_String.caseInsensitiveContains(workItem.fields["System.AreaPath"] || "", filterText)
                    || Utils_String.caseInsensitiveContains(status, filterText);
            });
        }
    }

    @autobind
    private _showDiscussions(workItem: WorkItem): void {
        this.setState(this._mergeState({showDiscussions: true, discussionWorkItem: workItem}));
    }

    @autobind
    private _hideDiscussions(): void {
        this.setState(this._mergeState({showDiscussions: false, discussionWorkItem: null}));
    }

    @autobind
    private _updateFilterText(searchText: string): void {
        this.setState(this._mergeState({filterText: searchText}));
    }

    @autobind
    private _changeSort(sortColumn: string, sortOrder: string): void {
        this.setState(this._mergeState({sortColumn: sortColumn, sortOrder: sortOrder}));
    }

    @autobind
    private _getMenuItems(): IContextualMenuItem[] {
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
                    this._refreshWorkItemResults();
                }
            },
            {
                key: "OpenQuery", name: "Open as query", title: "Open all workitems as a query", iconProps: {iconName: "OpenInNewWindow"}, 
                disabled: this.state.workItemResults.length === 0,
                onClick: async (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                    let url = `${VSS.getWebContext().host.uri}/${VSS.getWebContext().project.id}/_workitems?_a=query&wiql=${encodeURIComponent(this._getWiql().query)}`;
                    window.open(url, "_parent");
                }
            },
            {
                key: "Unlink", name: "Unlink workitems", title: "Unlink all workitems from the bug bash instance", iconProps: {iconName: "RemoveLink"}, 
                disabled: this.state.workItemResults.length === 0,
                onClick: async (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                    this._removeWorkItemsFromBugBash();
                }
            }
        ];        
    }

    private async _refreshWorkItemResults() {
        this.setState(this._mergeState({resultsLoading: true, resultsLoaded: false, workItemResults: [], filterText: "", showDiscussions: false, discussionWorkItem: null}));

        let queryResult = await WitClient.getClient().queryByWiql(this._getWiql(), VSS.getWebContext().project.id);
        let workItemIds = queryResult.workItems.map(workItem => workItem.id);
        let workItems: WorkItem[];

        if (workItemIds.length > 0) {
            workItems = await WitClient.getClient().getWorkItems(workItemIds);
        }
        else {
            workItems = [];
        }

        this.setState(this._mergeState({resultsLoading: false, resultsLoaded: true, workItemResults: workItems, filterText: "", showDiscussions: false, discussionWorkItem: null}));
    }

    private _mergeState(newState: any) {
        let currentState = this.state || this._getDefaultState();
        return {...currentState, ...newState};
    }

    private _getWiql(): Wiql {
        const item = this.props.context.stores.bugBashItemStore.getItem(this.props.id);

        return {
            query: `SELECT [System.Id], [System.Title], [System.CreatedBy], [System.CreatedDate], [System.State], [System.AssignedTo], [System.AreaPath] 
                    FROM WorkItems 
                    WHERE [System.TeamProject] = @project 
                    AND [System.WorkItemType] = '${item.workItemType}' 
                    AND [System.Tags] CONTAINS '${Helpers.getBugBashTag(item.id)}' 
                    ORDER BY [System.CreatedDate] DESC`
        };
    }

    @autobind
    private _refreshWorkItemsView(workItems: WorkItem[]) {
        this.setState(this._mergeState({workItemResults: workItems}));
    }

    @autobind
    private _addOrEditWorkItemsInView(workItems: WorkItem[]) {
        let newWorkItemResults: WorkItem[] = this.state.workItemResults.slice();

        for (const workItem of workItems) {
            let index = Utils_Array.findIndex(newWorkItemResults, (w: WorkItem) => w.id === workItem.id);
            if (index == -1) {
                newWorkItemResults.push(workItem);
            }
            else {
                newWorkItemResults[index] = workItem;
            }
        }
        this.setState(this._mergeState({workItemResults: newWorkItemResults}));
    }

    @autobind
    private _addWorkItem(workItem: WorkItem) {
        let workItems = [workItem].concat(this.state.workItemResults);
        this.setState(this._mergeState({workItemResults: workItems}));
    }

    @autobind
    private async _removeWorkItemsFromBugBash() {
        if (this.state.workItemResults.length > 0) {
            let dialogService: IHostDialogService = await VSS.getService(VSS.ServiceIds.Dialog) as IHostDialogService;
            try {
                await dialogService.openMessageDialog("Are you sure you want to remove all the work items from this bugbash instance?", { useBowtieStyle: true });  
                try {
                    Helpers.removeFromBugBash(this.state.item.id, this.state.workItemResults);                    
                    this.setState({...this.state, workItemError: null, workItemResults: []});
                }
                catch (e) {
                    this.setState({...this.state, workItemError: e.message});
                }          
            }
            catch (e) {
                // user selected "No"" in dialog
                return;
            }    
        }
    }
}