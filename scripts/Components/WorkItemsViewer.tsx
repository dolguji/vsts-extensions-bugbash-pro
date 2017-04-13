import * as React from "react";
import { DetailsList } from "OfficeFabric/DetailsList";
import { DetailsListLayoutMode, IColumn, CheckboxVisibility, ConstrainMode } from "OfficeFabric/components/DetailsList/DetailsList.Props";
import { SelectionMode } from "OfficeFabric/utilities/selection/interfaces";
import { Selection } from "OfficeFabric/utilities/selection/Selection";
import { autobind } from "OfficeFabric/Utilities";
import { IContextualMenuItem } from "OfficeFabric/components/ContextualMenu/ContextualMenu.Props";
import { IconButton } from "OfficeFabric/Button";
import { ContextualMenu } from "OfficeFabric/ContextualMenu";

import { WorkItem } from "TFS/WorkItemTracking/Contracts";
import * as WitBatchClient from "TFS/WorkItemTracking/BatchRestClient";
import { WorkItemFormNavigationService } from "TFS/WorkItemTracking/Services";
import Utils_Date = require("VSS/Utils/Date");
import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");

import { IBaseProps, IBugBash, Constants } from "../Models";
import { Loading } from "./Loading";
import { MessagePanel, MessageType } from "./MessagePanel";
import { IdentityView } from "./IdentityView";
import Helpers = require("../Helpers");

interface IWorkItemsViewerState {
    workItemError?: string;
    isContextMenuVisible?: boolean;
    contextMenuTarget?: MouseEvent;
}

export interface IWorkItemsViewerProps extends IBaseProps {
    areResultsReady: boolean;
    workItems: WorkItem[];    
    sortColumn: string;
    sortOrder: string;    
    bugBashItem: IBugBash;
    refreshWorkItems: (workItems: WorkItem[]) => void;
    changeSort: (sortColumn: string, sortOrder: string) => void;
    onShowDiscussions: (workItem: WorkItem) => void;
    updateWorkItems: (workItems: WorkItem[]) => void;
}

export class WorkItemsViewer extends React.Component<IWorkItemsViewerProps, IWorkItemsViewerState> {
    private _selection: Selection;

    constructor(props: IWorkItemsViewerProps, context: any) {
        super(props, context);

        this._selection = new Selection();
        this.state = {
            workItemError: null,
            isContextMenuVisible: false,
            contextMenuTarget: null    
        };
    }

    public componentWillReceiveProps(nextProps: IWorkItemsViewerProps) {
        this.setState({...this.state, workItemError: null});
    }

    public render(): JSX.Element {
        if (!this.props.areResultsReady) {
            return <Loading />;
        }
        else {
            let workItemsMenuItems: IContextualMenuItem[] = [
                {
                    key: "Delete", name: "Delete", title: "Delete selected workitems", iconProps: {iconName: "Delete"}, 
                    disabled: this._selection.getSelectedCount() == 0,
                    onClick: (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                        this._deleteSelectedWorkItems();
                    }
                },
                {
                    key: "Accept", name: "Accept", title: "Accept selected workitems", iconProps: {iconName: "SkypeCircleCheck"}, 
                    disabled: this._selection.getSelectedCount() == 0,
                    onClick: (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                        this._acceptOrRejectWorkItems(this._selection.getSelection() as WorkItem[], true);
                    }
                },
                {
                    key: "Reject", name: "Reject", title: "Reject selected workitems", iconProps: {iconName: "StatusErrorFull"}, 
                    disabled: this._selection.getSelectedCount() == 0,
                    onClick: (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                        this._acceptOrRejectWorkItems(this._selection.getSelection() as WorkItem[], false);
                    }
                },
                {
                    key: "OpenQuery", name: "Open as query", title: "Open selected workitems as a query", iconProps: {iconName: "OpenInNewWindow"}, 
                    disabled: this._selection.getSelectedCount() == 0,
                    onClick: (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                        let url = `${VSS.getWebContext().host.uri}/${VSS.getWebContext().project.id}/_workitems?_a=query&wiql=${encodeURIComponent(this._getSelectedWorkItemsWiql())}`;
                        window.open(url, "_parent");
                    }
                },
                {
                    key: "Remove", name: "Unlink from bug bash", title: "Unlink selected workitems from the bug bash instance", iconProps: {iconName: "RemoveLink"}, 
                    disabled: this._selection.getSelectedCount() == 0,
                    onClick: (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                        this._removeSelectedWorkItemsFromBugBash();
                    }
                }
            ];

            return (                
                <div className="results-view-contents">
                    { this.state.workItemError && <MessagePanel message={this.state.workItemError} messageType={MessageType.Error} />}
                    { this.props.workItems.length == 0 && (<MessagePanel message="No work item results." messageType={MessageType.Info} />) }
                    { this.props.workItems.length > 0 && (
                        <DetailsList 
                            layoutMode={DetailsListLayoutMode.justified}
                            constrainMode={ConstrainMode.horizontalConstrained}
                            selectionMode={SelectionMode.multiple}
                            isHeaderVisible={true}
                            checkboxVisibility={CheckboxVisibility.onHover}
                            columns={this._getColumns()}
                            onRenderItemColumn={this._onRenderCell}
                            items={this.props.workItems}
                            className="workitem-list"
                            onItemInvoked={(item: WorkItem, index: number) => {
                                this._openWorkItemDialog(null, item);
                            }}
                            selection={ this._selection }
                            onItemContextMenu={this._showContextMenu}
                            onColumnHeaderClick={this._onColumnHeaderClick}
                        />                                    
                    ) }
                    { this.state.isContextMenuVisible && (
                        <ContextualMenu
                            className="context-menu"
                            items={workItemsMenuItems}
                            target={this.state.contextMenuTarget}
                            shouldFocusOnMount={ true }
                            onDismiss={this._hideContextMenu}
                        />
                    )}
                </div>                
            );
        }
    }

    private _getColumns(): IColumn[] {
        return [
            {
                fieldName: Constants.ACCEPT_STATUS_CELL_NAME,
                key: "Status",
                name:"Status",
                minWidth: 40,
                maxWidth: 70,
                isResizable: true,
                isSorted: Utils_String.equals(this.props.sortColumn, Constants.ACCEPT_STATUS_CELL_NAME, true),
                isSortedDescending: Utils_String.equals(this.props.sortOrder, "desc", true)
            },
            {
                fieldName: "ID",
                key: "ID",
                name:"ID",
                minWidth: 40,
                maxWidth: 70,
                isResizable: true,
                isSorted: Utils_String.equals(this.props.sortColumn, "ID", true),
                isSortedDescending: Utils_String.equals(this.props.sortOrder, "desc", true)
            },
            {
                fieldName: "System.Title",
                key: "Title",
                name:"Title",
                minWidth: 150,
                maxWidth: 300,
                isResizable: true,
                isSorted: Utils_String.equals(this.props.sortColumn, "System.Title", true),
                isSortedDescending: Utils_String.equals(this.props.sortOrder, "desc", true)
            },
            {
                fieldName: "System.CreatedBy",
                key: "CreatedBy",
                name:"Created By",
                minWidth: 100,
                maxWidth: 250,
                isResizable: true,
                isSorted: Utils_String.equals(this.props.sortColumn, "System.CreatedBy", true),
                isSortedDescending: Utils_String.equals(this.props.sortOrder, "desc", true)
            },
            {
                fieldName: "System.CreatedDate",
                key: "CreatedDate",
                name:"Created Date",
                minWidth: 120,
                maxWidth: 170,
                isResizable: true,
                isSorted: Utils_String.equals(this.props.sortColumn, "System.CreatedDate", true),
                isSortedDescending: Utils_String.equals(this.props.sortOrder, "desc", true)
            },
            {
                fieldName: "System.State",
                key: "State",
                name:"State",
                minWidth: 100,
                maxWidth: 150,
                isResizable: true,
                isSorted: Utils_String.equals(this.props.sortColumn, "System.State", true),
                isSortedDescending: Utils_String.equals(this.props.sortOrder, "desc", true)
            },
            {
                fieldName: "System.AssignedTo",
                key: "AssignedTo",
                name:"Assigned To",
                minWidth: 100,
                maxWidth: 250,
                isResizable: true,
                isSorted: Utils_String.equals(this.props.sortColumn, "System.AssignedTo", true),
                isSortedDescending: Utils_String.equals(this.props.sortOrder, "desc", true)
            },
            {
                fieldName: "System.AreaPath",
                key: "AreaPath",
                name:"Area Path",
                minWidth: 150,
                maxWidth: 350,
                isResizable: true,
                isSorted: Utils_String.equals(this.props.sortColumn, "System.AreaPath", true),
                isSortedDescending: Utils_String.equals(this.props.sortOrder, "desc", true)
            },
            {
                fieldName: Constants.ACTIONS_CELL_NAME,
                key: Constants.ACTIONS_CELL_NAME,
                name: "",
                minWidth: 120,
                maxWidth: 120,
                isResizable: true,
                isSorted: false,
                isSortedDescending: false
            }
        ];
    }    

    @autobind
    private _onColumnHeaderClick(ev?: React.MouseEvent<HTMLElement>, column?: IColumn) {
        if (column.fieldName !== Constants.ACTIONS_CELL_NAME) {
            this.props.changeSort(column.fieldName, column.isSortedDescending ? "asc" : "desc");
        }
    }

    @autobind
    private _onRenderCell(item: WorkItem, index?: number, column?: IColumn): React.ReactNode {
        let text: string;
        switch (column.fieldName) {
            case Constants.ACCEPT_STATUS_CELL_NAME:
                let classNames = "overflow-ellipsis";
                let statusText = "";
                if (Helpers.isWorkItemAccepted(item)) {
                    classNames += " workitem-accepted";
                    statusText = Constants.ACCEPTED_TEXT;
                }
                else if(Helpers.isWorkItemRejected(item)) {
                    classNames += " workitem-rejected";
                    statusText = Constants.REJECTED_TEXT;
                }
                return <span className={classNames}>{statusText}</span>;
            case "ID":
                text = `${item.id}`;
                break;
            case "System.Title":
                return <span className="title-cell overflow-ellipsis" onClick={(e) => this._openWorkItemDialog(e, item)} title={item.fields[column.fieldName]}>{item.fields[column.fieldName]}</span>
            case "System.CreatedDate":
                text = Utils_Date.friendly(new Date(item.fields["System.CreatedDate"]));
                break;
            case "System.CreatedBy":
            case "System.AssignedTo":
                return <IdentityView identityDistinctName={item.fields[column.fieldName]} />;
            case Constants.ACTIONS_CELL_NAME:                
                return (
                    <div className="workitem-row-actions-cell">
                        <IconButton icon="Chat" className="workitem-row-cell-button" title="Show discussions" onClick={() => this.props.onShowDiscussions(item)} />
                        {!Helpers.isWorkItemAccepted(item) && <IconButton className="accept-button workitem-row-cell-button" icon="SkypeCircleCheck" title="Accept workitem" onClick={() => this._acceptOrRejectWorkItems([item], true)} />}
                        {!Helpers.isWorkItemRejected(item) && <IconButton className="reject-button workitem-row-cell-button" icon="StatusErrorFull" title="Reject workitem" onClick={() => this._acceptOrRejectWorkItems([item], false)} />}
                    </div>
                );
            default:
                text = item.fields[column.fieldName];  
                break;          
        }

        return <div className="overflow-ellipsis" title={text}>{text}</div>;
    }    

    @autobind
    private _showContextMenu(item?: WorkItem, index?: number, e?: MouseEvent) {
        if (!this._selection.isIndexSelected(index)) {
            // if not already selected, unselect every other row and select this one
            this._selection.setAllSelected(false);
            this._selection.setIndexSelected(index, true, true);
        }        
        this.setState({...this.state, contextMenuTarget: e, isContextMenuVisible: true});
    }

    @autobind
    private _hideContextMenu(e?: any) {
        this.setState({...this.state, contextMenuTarget: null, isContextMenuVisible: false});
    }

    private async _acceptOrRejectWorkItems(workItems: WorkItem[], accept: boolean) {
        let templateFieldValues: IDictionaryStringTo<string> = {};
        let fieldValuesMap: IDictionaryNumberTo<IDictionaryStringTo<string>> = {};
        const configTemplateKey = accept ? Constants.ACCEPT_CONFIG_TEMPLATE_KEY : Constants.REJECT_CONFIG_TEMPLATE_KEY;
        const tagToBeAdded = accept ? Constants.BUGBASH_ACCEPT_TAG : Constants.BUGBASH_REJECT_TAG;
        const tagToBeRemoved = accept ? Constants.BUGBASH_REJECT_TAG : Constants.BUGBASH_ACCEPT_TAG;

        if (this.props.bugBashItem.configTemplates && this.props.bugBashItem.configTemplates[configTemplateKey]) {
            let templateFound = await this.props.context.actionsCreator.ensureTemplateItem(this.props.bugBashItem.configTemplates[configTemplateKey]);

            if (templateFound) {
                let template = this.props.context.stores.workItemTemplateItemStore.getItem(this.props.bugBashItem.configTemplates[configTemplateKey]);
                templateFieldValues = { ...template.fields };
            }
        }

        for (const workItem of workItems) {
            let fieldValues: IDictionaryStringTo<string> = {...templateFieldValues};

            // add Accept/Reject tag, bug bash tag and remove accept/reject tag
            let tagArr = Helpers.parseTags(workItem.fields["System.Tags"]);
            tagArr.push(tagToBeAdded, Helpers.getBugBashTag(this.props.bugBashItem.id));
            let rejectedTagIndex = tagArr.indexOf(tagToBeRemoved);
            if (rejectedTagIndex !== -1) {
                tagArr.splice(rejectedTagIndex, 1);
            }
            
            tagArr = tagArr.concat(Helpers.parseTags(fieldValues["System.Tags-Add"]));
            tagArr = Utils_Array.subtract(tagArr, Helpers.parseTags(fieldValues["System.Tags-Remove"]), Utils_String.ignoreCaseComparer);

            delete fieldValues["System.Tags-Add"];
            delete fieldValues["System.Tags-Remove"];
            delete fieldValues["System.Tags"];
            
            fieldValues["System.Tags"] = tagArr.join(";");

            fieldValuesMap[workItem.id] = fieldValues;
        }
    
        try {
            let workItems = await Helpers.saveWorkItems(fieldValuesMap);
            this.setState({...this.state, workItemError: null});
            this.props.updateWorkItems(workItems);
        }
        catch (e) {
            this.setState({...this.state, workItemError: e.message});
        }
    }    

    @autobind
    private async _openWorkItemDialog(e: React.MouseEvent<HTMLElement>, item: WorkItem) {
        let newTab = e ? e.ctrlKey : false;
        let workItemNavSvc = await WorkItemFormNavigationService.getService();
        workItemNavSvc.openWorkItem(item.id, newTab);
    }

    @autobind
    private async _removeSelectedWorkItemsFromBugBash() {
        let selectedWorkItems = this._selection.getSelection() as WorkItem[];
        if (selectedWorkItems.length > 0) {
            let dialogService: IHostDialogService = await VSS.getService(VSS.ServiceIds.Dialog) as IHostDialogService;
            try {
                await dialogService.openMessageDialog("Are you sure you want to remove the selected work items from this bugbash instance?", { useBowtieStyle: true });  
                try {
                    Helpers.removeFromBugBash(this.props.bugBashItem.id, selectedWorkItems);
                    
                    let workItemResults = this.props.workItems.filter((item: WorkItem) => {
                        return Utils_Array.findIndex(selectedWorkItems, (wi: WorkItem) => Utils_String.equals(""+wi.id, ""+item.id)) === -1;
                    });

                    this._selection.setAllSelected(false);
                    this.setState({...this.state, workItemError: null});
                    this.props.refreshWorkItems(workItemResults);
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

    @autobind
    private async _deleteSelectedWorkItems() {
        let selectedWorkItems = this._selection.getSelection() as WorkItem[];
        if (selectedWorkItems.length > 0) {
            let dialogService: IHostDialogService = await VSS.getService(VSS.ServiceIds.Dialog) as IHostDialogService;
            try {
                await dialogService.openMessageDialog("Are you sure you want to delete the selected work items?", { useBowtieStyle: true });  
                try {
                    await WitBatchClient.getClient().deleteWorkItemsBatch(selectedWorkItems.map((item: WorkItem) => item.id));
                    let workItemResults = this.props.workItems.filter((item: WorkItem) => {
                        return Utils_Array.findIndex(selectedWorkItems, (wi: WorkItem) => Utils_String.equals(""+wi.id, ""+item.id)) === -1;
                    });

                    this._selection.setAllSelected(false);
                    this.setState({...this.state, workItemError: null});
                    this.props.refreshWorkItems(workItemResults);
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

    private _getSelectedWorkItemsWiql(): string {
        let selectedWorkItems = this._selection.getSelection() as WorkItem[];
        let ids = selectedWorkItems.map((w:WorkItem) => w.id).join(",");

        return `SELECT [System.Id], [System.Title], [System.CreatedBy], [System.CreatedDate], [System.State], [System.AssignedTo], [System.AreaPath]
                 FROM WorkItems 
                 WHERE [System.TeamProject] = @project 
                 AND [System.ID] IN (${ids}) 
                 ORDER BY [System.CreatedDate] DESC`;
    }
}