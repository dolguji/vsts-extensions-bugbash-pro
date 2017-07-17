import "../../css/AllBugBashes.scss";

import * as React from "react";

import { SelectionMode } from "OfficeFabric/utilities/selection";
import { TooltipHost, TooltipDelay, DirectionalHint, TooltipOverflowMode } from "OfficeFabric/Tooltip";
import { Label } from "OfficeFabric/Label";
import { IContextualMenuItem } from "OfficeFabric/components/ContextualMenu/ContextualMenu.Props";
import { Panel, PanelType } from "OfficeFabric/Panel";

import { MessagePanel, MessageType } from "VSTS_Extension/Components/Common/MessagePanel";
import { Loading } from "VSTS_Extension/Components/Common/Loading";
import { Grid } from "VSTS_Extension/Components/Grids/Grid";
import { QueryResultGrid } from "VSTS_Extension/Components/Grids/WorkItemGrid/QueryResultGrid";
import { IContextMenuProps, GridColumn } from "VSTS_Extension/Components/Grids/Grid.Props";
import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { BaseStore } from "VSTS_Extension/Flux/Stores/BaseStore";
import { LazyLoad } from "VSTS_Extension/Components/Common/LazyLoad";
import { Hub } from "VSTS_Extension/Components/Common/Hub/Hub";

import { HostNavigationService } from "VSS/SDK/Services/Navigation";
import Utils_Date = require("VSS/Utils/Date");
import Utils_String = require("VSS/Utils/String");
import Context = require("VSS/Context");

import { confirmAction } from "../Helpers";
import { UrlActions } from "../Constants";
import { IBugBash } from "../Interfaces";
import { StoresHub } from "../Stores/StoresHub";
import { BugBashActions } from "../Actions/BugBashActions";

interface IAllBugBashesViewState extends IBaseComponentState {
    loading: boolean,    
    pastBugBashes: IBugBash[];
    currentBugBashes: IBugBash[];
    upcomingBugBashes: IBugBash[];
    settingsPanelOpen: boolean;
    selectedPivot?: string;
}

export class AllBugBashesView extends BaseComponent<IBaseComponentProps, IAllBugBashesViewState> {
    protected getStores(): BaseStore<any, any, any>[] {
        return [StoresHub.bugBashStore];
    }

    protected initializeState() {
        this.state = {
            pastBugBashes: [],
            currentBugBashes: [],
            upcomingBugBashes: [],
            loading: true,
            selectedPivot: "ongoing",
            settingsPanelOpen: false
        };
    }

    public componentDidMount() {
        super.componentDidMount();
        BugBashActions.initializeAllBugBashes(); 
    }

    protected getStoresState(): IAllBugBashesViewState {        
        let allBugBashes = StoresHub.bugBashStore.getAll() || [];
        let currentTime = new Date();
        allBugBashes = allBugBashes.filter((bugBash: IBugBash) => Utils_String.equals(VSS.getWebContext().project.id, bugBash.projectId, true));
        
        return {
            pastBugBashes: this._getPastBugBashes(allBugBashes, currentTime),
            currentBugBashes: this._getCurrentBugBashes(allBugBashes, currentTime),
            upcomingBugBashes: this._getUpcomingBugBashes(allBugBashes, currentTime),
            loading: StoresHub.bugBashStore.isLoading()
        } as IAllBugBashesViewState;
    }

    public render(): JSX.Element {
        return (
            <div className="all-view">
                <Hub 
                    title="Bug Bashes"          
                    pivotProps={{
                        initialSelectedKey: this.state.selectedPivot,
                        onPivotClick: (selectedPivotKey: string, ev?: React.MouseEvent<HTMLElement>) => {
                            this.updateState({selectedPivot: selectedPivotKey} as IAllBugBashesViewState);
                        },
                        onRenderPivotContent: (key: string) => {
                            return this._getContents(key);
                        },
                        pivots: [
                            {
                                key: "ongoing",
                                text: "Ongoing",
                                itemCount: this.state.currentBugBashes ? this.state.currentBugBashes.length : null,
                                commands: this._getCommandBarItems()
                            },
                            {
                                key: "upcoming",
                                text: "UpComing",
                                itemCount: this.state.upcomingBugBashes ? this.state.upcomingBugBashes.length : null,
                                commands: this._getCommandBarItems()
                            },
                            {
                                key: "past",
                                text: "Past",
                                itemCount: this.state.pastBugBashes ? this.state.pastBugBashes.length : null,
                                commands: this._getCommandBarItems()
                            }                        
                        ]
                    }}
                />                

                { 
                    this.state.settingsPanelOpen && 
                    <Panel
                        isOpen={true}
                        type={PanelType.smallFixedFar}
                        isLightDismiss={true} 
                        onDismiss={() => this.updateState({settingsPanelOpen: false} as IAllBugBashesViewState)}>

                        <LazyLoad module="scripts/SettingsPanel">
                            {(SettingsPanel) => (
                                <SettingsPanel.SettingsPanel />
                            )}
                        </LazyLoad>
                    </Panel>
                }
            </div>
        );
    }

    private _getContents(key: string): JSX.Element {
        if (this.state.loading) {
            return <Loading />;
        }
            
        let bugBashes: IBugBash[];
        let missingItemsMsg = "";

        if (key === "past") {
            bugBashes = this.state.pastBugBashes;
            missingItemsMsg = "No past bug bashes.";
        }
        else if (key === "ongoing") {
            bugBashes = this.state.currentBugBashes;
            missingItemsMsg = "No ongoing bug bashes.";
        }
        else {
            bugBashes = this.state.upcomingBugBashes;
            missingItemsMsg = "No upcoming bug bashes.";            
        }

        return <Grid
            className={"instance-list"}
            items={bugBashes}
            columns={this._getGridColumns()}
            selectionMode={SelectionMode.none}
            contextMenuProps={this._getGridContextMenuProps()}
            noResultsText={missingItemsMsg}
        />;                    
    }

    private _getGridColumns(): GridColumn[] {
        return [
            {
                key: "title",
                name: "Title",
                minWidth: 300,
                maxWidth: Infinity,
                onRenderCell: bugBash => {
                    return <TooltipHost 
                        content={bugBash.title}
                        delay={TooltipDelay.medium}
                        overflowMode={TooltipOverflowMode.Parent}
                        directionalHint={DirectionalHint.bottomLeftEdge}>
                        <Label className="bugbash-grid-cell">
                            <a href={this._getBugBashUrl(bugBash, UrlActions.ACTION_RESULTS)} onClick={(e: React.MouseEvent<HTMLElement>) => this._onRowClick(e, bugBash)}>{ bugBash.title }</a>
                        </Label>
                    </TooltipHost>;
                }
            },
            {
                key: "startDate",
                name: "Start Date",
                minWidth: 400,
                maxWidth: 600,
                onRenderCell: bugBash => {
                    const label = bugBash.startTime ? Utils_Date.format(bugBash.startTime, "dddd, MMMM dd, yyyy") : "N/A";
                    return <TooltipHost 
                        content={label}
                        delay={TooltipDelay.medium}
                        overflowMode={TooltipOverflowMode.Parent}
                        directionalHint={DirectionalHint.bottomLeftEdge}>
                        <Label className="bugbash-grid-cell">{label}</Label>
                    </TooltipHost>;
                }
            },
            {
                key: "endDate",
                name: "End Date",
                minWidth: 400,
                maxWidth: 600,
                onRenderCell: bugBash => {
                    const label = bugBash.endTime ? Utils_Date.format(bugBash.endTime, "dddd, MMMM dd, yyyy") : "N/A";
                    return <TooltipHost 
                        content={label}
                        delay={TooltipDelay.medium}
                        overflowMode={TooltipOverflowMode.Parent}
                        directionalHint={DirectionalHint.bottomLeftEdge}>
                        <Label className="bugbash-grid-cell">{label}</Label>
                    </TooltipHost>;
                }
            }
        ];
    }

    private _getGridContextMenuProps(): IContextMenuProps {
        return {
            menuItems: (selectedItems: IBugBash[]) => {
                if (selectedItems.length !== 1) {
                    return [];
                }

                const bugBash = selectedItems[0];
                return [
                    {
                        key: "open", name: "View results", iconProps: {iconName: "ShowResults"}, 
                        onClick: async (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {                    
                            let navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
                            navigationService.updateHistoryEntry(UrlActions.ACTION_RESULTS, {id: bugBash.id});
                        }
                    },
                    {
                        key: "edit", name: "Edit", iconProps: {iconName: "Edit"}, 
                        onClick: async (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {                    
                            let navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
                            navigationService.updateHistoryEntry(UrlActions.ACTION_EDIT, {id: bugBash.id});
                        }
                    },
                    {
                        key: "delete", name: "Delete", iconProps: {iconName: "Cancel", style: { color: "#da0a00", fontWeight: "bold" }}, 
                        onClick: async (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {                    
                            const confirm = await confirmAction(true, "Are you sure you want to delete this bug bash instance?");
                            if (confirm) {
                                BugBashActions.deleteBugBash(bugBash.id);                                
                            }   
                        }
                    },
                ];
            }
        }
    }

    private _getCommandBarItems(): IContextualMenuItem[] {
         return [
            {
                key: "new", name: "New", iconProps: {iconName: "Add"},
                onClick: async (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                    let navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
                    navigationService.updateHistoryEntry(UrlActions.ACTION_EDIT);
                }
            },            
            {
                key: "refresh", name: "Refresh", iconProps: {iconName: "Refresh"},
                onClick: async (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                    BugBashActions.refreshAllBugBashes();
                }
            },
            {
                key: "settings", name: "Settings", iconProps: {iconName: "Settings"},
                onClick: async (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                    this.updateState({settingsPanelOpen: !(this.state.settingsPanelOpen)} as IAllBugBashesViewState);
                }
            }
         ];
    }

    private _getBugBashUrl(bugBash: IBugBash, viewName?: string): string {
        const pageContext = Context.getPageContext();
        const navigation = pageContext.navigation;
        const webContext = VSS.getWebContext();
        return `${webContext.collection.uri}/${webContext.project.name}/_${navigation.currentController}/${navigation.currentAction}/${navigation.currentParameters}#_a=${viewName}&id=${bugBash.id}`;
    }

    private async _onRowClick(e: React.MouseEvent<HTMLElement>, bugBash: IBugBash) {
        if (!e.ctrlKey) {
            e.preventDefault();
            let navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
            navigationService.updateHistoryEntry(UrlActions.ACTION_RESULTS, {id: bugBash.id});
        }      
    }

    private _getPastBugBashes(list: IBugBash[], currentTime: Date): IBugBash[] {
        return list.filter((bugBash: IBugBash) => {
            return bugBash.endTime && Utils_Date.defaultComparer(bugBash.endTime, currentTime) < 0;
        }).sort((b1: IBugBash, b2: IBugBash) => {
            return Utils_Date.defaultComparer(b1.endTime, b2.endTime);
        });
    }

    private _getCurrentBugBashes(list: IBugBash[], currentTime: Date): IBugBash[] {        
        return list.filter((bugBash: IBugBash) => {
            if (!bugBash.startTime && !bugBash.endTime) {
                return true;
            }
            else if(!bugBash.startTime && bugBash.endTime) {
                return Utils_Date.defaultComparer(bugBash.endTime, currentTime) >= 0;
            }
            else if (bugBash.startTime && !bugBash.endTime) {
                return Utils_Date.defaultComparer(bugBash.startTime, currentTime) <= 0;
            }
            else {
                return Utils_Date.defaultComparer(bugBash.startTime, currentTime) <= 0 && Utils_Date.defaultComparer(bugBash.endTime, currentTime) >= 0;
            }
        }).sort((b1: IBugBash, b2: IBugBash) => {
            return Utils_Date.defaultComparer(b1.startTime, b2.startTime);
        });
    }

    private _getUpcomingBugBashes(list: IBugBash[], currentTime: Date): IBugBash[] {
        return list.filter((bugBash: IBugBash) => {
            return bugBash.startTime && Utils_Date.defaultComparer(bugBash.startTime, currentTime) > 0;
        }).sort((b1: IBugBash, b2: IBugBash) => {
            return Utils_Date.defaultComparer(b1.startTime, b2.startTime);
        });
    }
}