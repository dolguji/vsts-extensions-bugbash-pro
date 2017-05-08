import "../../css/ResultsView.scss";

import * as React from "react";

import { HostNavigationService } from "VSS/SDK/Services/Navigation";
import * as WitClient from "TFS/WorkItemTracking/RestClient";
import { Wiql, WorkItem } from "TFS/WorkItemTracking/Contracts";
import Utils_Date = require("VSS/Utils/Date");
import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");

import { MessageBar, MessageBarType } from "OfficeFabric/MessageBar";
import { Label } from "OfficeFabric/Label";
import { IContextualMenuItem } from "OfficeFabric/components/ContextualMenu/ContextualMenu.Props";
import { autobind } from "OfficeFabric/Utilities";

import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { BaseStore } from "VSTS_Extension/Stores/BaseStore";
import { WorkItemFieldStore } from "VSTS_Extension/Stores/WorkItemFieldStore";
import { Loading } from "VSTS_Extension/Components/Common/Loading";
import { Grid } from "VSTS_Extension/Components/Grids/Grid";
import { SortOrder, GridColumn, ICommandBarProps, IContextMenuProps } from "VSTS_Extension/Components/Grids/Grid.Props";

import { IBugBash, UrlActions, IBugBashItemDocument } from "../Models";
import { BugBashItem } from "./BugBashItem";
import Helpers = require("../Helpers");
import { BugBashStore } from "../Stores/BugBashStore";
import { StoresHub } from "../Stores/StoresHub";

interface IViewHubViewState extends IBaseComponentState {
    bugBashItem?: IBugBash;
    items?: IBugBashItemDocument[];
    selectedBugBashItem?: IBugBashItemDocument;
    loading?: boolean;
}

interface IViewHubViewProps extends IBaseComponentProps {
    id: string;
}

export class ViewBugBashView extends BaseComponent<IViewHubViewProps, IViewHubViewState> {
    protected getStoresToLoad(): {new (): BaseStore<any, any, any>}[] {
        return [BugBashStore];
    }

    protected initializeState() {
        this.state = {
            bugBashItem: null,
            loading: true,
            items: [],
            selectedBugBashItem: null   
        };
    }

    protected async initialize() {                
        let found = await StoresHub.bugBashStore.ensureItem(this.props.id);
        if (!found) {
            this.updateState({
                bugBashItem: null,
                loading: false,
                items: [],
                selectedBugBashItem: null
            });
        }
        else {
            //this._refreshWorkItemResults();
        }
    }   

    protected onStoreChanged() {
        this.updateState({bugBashItem: StoresHub.bugBashStore.getItem(this.props.id), loading: !StoresHub.bugBashStore.isLoaded()});
    }

    public render(): JSX.Element {
        if (this.state.loading) {
            return <Loading />;
        }
        else {
            if (!this.state.bugBashItem) {
                return <MessageBar messageBarType={MessageBarType.error}>This instance of bug bash either doesn't exist or is out of scope of current project.</MessageBar>;
            }
            else {
                return (
                    <div className="results-view">                        
                        <Grid
                            className="bugbash-item-grid"
                            items={this.state.items}
                            columns={this._getGridColumns()}
                            commandBarProps={{menuItems: this._getCommandBarMenuItems(), farMenuItems: this._getCommandBarFarMenuItems()}}
                            contextMenuProps={{menuItems: this._getContextMenuItems}}
                            onItemInvoked={(item: IBugBashItemDocument) => this.updateState({selectedBugBashItem: item})}
                        />
                        
                        <div className="item-viewer">
                            <BugBashItem item={this.state.selectedBugBashItem} />
                        </div>
                    </div>
                );
            }            
        }
    }     

    @autobind
    private _showItem(item?: IBugBashItemDocument): void {
        this.updateState({selectedBugBashItem: item});
    }

    @autobind
    private _getGridColumns(): GridColumn[] {
        return [
            {
                key: "title",
                name: "Title",
                minWidth: 150,
                maxWidth: Infinity,
                resizable: true,
                onRenderCell: (item: IBugBashItemDocument) => <Label>{item.title}</Label>,
                sortFunction: (item1: IBugBashItemDocument, item2: IBugBashItemDocument, sortOrder: SortOrder) => Utils_String.ignoreCaseComparer(item1.title, item2.title),
                filterFunction: (item: IBugBashItemDocument, filterText: string) => Utils_String.caseInsensitiveContains(item.title, filterText)
            }
        ];
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
                        //this._refreshWorkItemResults();
                    }
                },
                {
                    key: "OpenQuery", name: "Open as query", title: "Open all workitems as a query", iconProps: {iconName: "OpenInNewWindow"}, 
                    disabled: false, //this.state.workItemResults.length === 0,
                    onClick: async (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                        //let url = `${VSS.getWebContext().host.uri}/${VSS.getWebContext().project.id}/_workitems?_a=query&wiql=${encodeURIComponent(this._getWiql().query)}`;
                        //window.open(url, "_parent");
                    }
                },
                {
                    key: "Unlink", name: "Unlink workitems", title: "Unlink all workitems from the bug bash instance", iconProps: {iconName: "RemoveLink"}, 
                    disabled: false, //this.state.workItemResults.length === 0,
                    onClick: async (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                        //this._removeWorkItemsFromBugBash();
                    }
                }
            ];
    }

    private _getCommandBarFarMenuItems(): IContextualMenuItem[] {
        return [
                {
                    key: "Home", name: "Home", title: "Return to home view", iconProps: {iconName: "Home"}, 
                    onClick: async (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                        let navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
                        navigationService.updateHistoryEntry(UrlActions.ACTION_ALL, null);
                    }
                }
            ];
    }

    @autobind
    private _getContextMenuItems(selectedItems: any[]): IContextualMenuItem[] {
        return [];
    }
}