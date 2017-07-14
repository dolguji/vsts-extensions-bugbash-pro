import "../../css/AllBugBashes.scss";

import * as React from "react";

import { List } from "OfficeFabric/List";
import { autobind } from "OfficeFabric/Utilities";
import { Label } from "OfficeFabric/Label";
import { CommandBar } from "OfficeFabric/CommandBar";
import { IContextualMenuItem } from "OfficeFabric/components/ContextualMenu/ContextualMenu.Props";
import { Panel, PanelType } from "OfficeFabric/Panel";

import { MessagePanel, MessageType } from "VSTS_Extension/Components/Common/MessagePanel";
import { Loading } from "VSTS_Extension/Components/Common/Loading";
import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { BaseStore } from "VSTS_Extension/Flux/Stores/BaseStore";
import { LazyLoad } from "VSTS_Extension/Components/Common/LazyLoad";

import { HostNavigationService } from "VSS/SDK/Services/Navigation";
import Utils_Date = require("VSS/Utils/Date");
import Utils_String = require("VSS/Utils/String");
import Context = require("VSS/Context");

import { UrlActions } from "../Constants";
import { IBugBash } from "../Interfaces";
import { StoresHub } from "../Stores/StoresHub";
import { BugBashActions } from "../Actions/BugBashActions";

interface IAllBugBashesViewState extends IBaseComponentState {
    loading?: boolean,
    allItems?: IBugBash[];
    pastItems?: IBugBash[];
    currentItems?: IBugBash[];
    upcomingItems?: IBugBash[];
    settingsPanelOpen?: boolean;
}

export class AllBugBashesView extends BaseComponent<IBaseComponentProps, IAllBugBashesViewState> {
    protected getStores(): BaseStore<any, any, any>[] {
        return [StoresHub.bugBashStore];
    }

    protected initializeState() {
        this.state = {
            allItems: [],
            pastItems: [],
            currentItems: [],
            upcomingItems: [],
            loading: true
        };
    }

    public componentDidMount() {
        super.componentDidMount();
        BugBashActions.initializeAllBugBashes(); 
    }

    protected getStoresState(): IAllBugBashesViewState {        
        let allItems = StoresHub.bugBashStore.getAll() || [];
        let currentTime = new Date();
        allItems = allItems.filter((item: IBugBash) => Utils_String.equals(VSS.getWebContext().project.id, item.projectId, true));
        
        return {
            allItems: allItems,
            pastItems: this._getPastBugBashes(allItems, currentTime),
            currentItems: this._getCurrentBugBashes(allItems, currentTime),
            upcomingItems: this._getUpcomingBugBashes(allItems, currentTime),
            loading: StoresHub.bugBashStore.isLoading()
        };
    }

    public render(): JSX.Element {
        return (
            <div className="all-view">
                <CommandBar className="all-view-menu-toolbar" items={this._getMenuItems()} />
                {this._getContents()}

                { 
                    this.state.settingsPanelOpen && 
                    <Panel
                        isOpen={true}
                        type={PanelType.smallFixedFar}
                        isLightDismiss={true} 
                        onDismiss={() => this.updateState({settingsPanelOpen: false})}>

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

    private _getContents(): JSX.Element {
        if (this.state.loading) {
            return <Loading />;
        }
        else {
            if (this.state.allItems.length == 0) {
                return <MessagePanel messageType={MessageType.Info} message="No instance of bug bash exists in the context of current project." />;
            }
            else {
                return (                    
                    <div className="instance-list-container">
                        <div className="instance-list-section">
                            <Label className="header">Past Bug Bashes ({this.state.pastItems.length})</Label>
                            <div className="instance-list-content">
                                {this.state.pastItems.length === 0 && <MessagePanel messageType={MessageType.Info} message="No past bug bashes." />}
                                {this.state.pastItems.length > 0 && <List items={this.state.pastItems} className="instance-list" onRenderCell={this._onRenderCell} />}
                            </div>
                        </div>

                        <div className="instance-list-section">
                            <Label className="header">Ongoing Bug Bashes ({this.state.currentItems.length})</Label>
                            <div className="instance-list-content">
                                {this.state.currentItems.length === 0 && <MessagePanel messageType={MessageType.Info} message="No ongoing bug bashes." />}
                                {this.state.currentItems.length > 0 && <List items={this.state.currentItems} className="instance-list" onRenderCell={this._onRenderCell} />}
                            </div>
                        </div>

                        <div className="instance-list-section">
                            <Label className="header">Upcoming Bug Bashes ({this.state.upcomingItems.length})</Label>
                            <div className="instance-list-content">
                                {this.state.upcomingItems.length === 0 && <MessagePanel messageType={MessageType.Info} message="No ongoing bug bashes." />}
                                {this.state.upcomingItems.length > 0 && <List items={this.state.upcomingItems} className="instance-list" onRenderCell={this._onRenderCell} />}
                            </div>
                        </div>
                    </div>
                );
            }
        }
    }    

    private _getMenuItems(): IContextualMenuItem[] {
         return [
            {
                key: "new", name: "New", title: "Create new instance", iconProps: {iconName: "Add"},
                onClick: async (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                    let navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
                    navigationService.updateHistoryEntry(UrlActions.ACTION_NEW);
                }
            },            
            {
                key: "refresh", name: "Refresh", title: "Refresh list", iconProps: {iconName: "Refresh"},
                onClick: async (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                    BugBashActions.refreshAllBugBashes();
                }
            },
            {
                key: "settings", name: "Settings", title: "Open settings panel", iconProps: {iconName: "Settings"},
                onClick: async (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                    this.updateState({settingsPanelOpen: !(this.state.settingsPanelOpen)});
                }
            }
         ];
    }

    @autobind
    private _onRenderCell(item: IBugBash, index?: number): React.ReactNode {
        const pageContext = Context.getPageContext();
        const navigation = pageContext.navigation;
        const webContext = VSS.getWebContext();
        const url = `${webContext.collection.uri}/${webContext.project.name}/_${navigation.currentController}/${navigation.currentAction}/${navigation.currentParameters}#_a=${UrlActions.ACTION_VIEW}&id=${item.id}`;

        return (
            <div className="instance-row">
                <a className="instance-title" href={url} onClick={(e: React.MouseEvent<HTMLElement>) => this._onRowClick(e, item)}>{ item.title }</a>
                <div className="instance-info">
                    <div className="instance-info-cell-container"><div className="instance-info-cell">Start:</div><div className="instance-info-cell-info">{item.startTime ? Utils_Date.format(item.startTime, "dddd, MMMM dd, yyyy") : "N/A"}</div></div>
                    <div className="instance-info-cell-container"><div className="instance-info-cell">End:</div><div className="instance-info-cell-info">{item.endTime ? Utils_Date.format(item.endTime, "dddd, MMMM dd, yyyy") : "N/A"}</div></div>
                </div>
            </div>
        );
    }

    private async _onRowClick(e: React.MouseEvent<HTMLElement>, item: IBugBash) {
        if (!e.ctrlKey) {
            e.preventDefault();
            let navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
            navigationService.updateHistoryEntry(UrlActions.ACTION_EDIT, {id: item.id});
        }      
    }

    private _getPastBugBashes(list: IBugBash[], currentTime: Date): IBugBash[] {
        return list.filter((item: IBugBash) => {
            return item.endTime && Utils_Date.defaultComparer(item.endTime, currentTime) < 0;
        }).sort((b1: IBugBash, b2: IBugBash) => {
            return Utils_Date.defaultComparer(b1.endTime, b2.endTime);
        });
    }

    private _getCurrentBugBashes(list: IBugBash[], currentTime: Date): IBugBash[] {        
        return list.filter((item: IBugBash) => {
            if (!item.startTime && !item.endTime) {
                return true;
            }
            else if(!item.startTime && item.endTime) {
                return Utils_Date.defaultComparer(item.endTime, currentTime) >= 0;
            }
            else if (item.startTime && !item.endTime) {
                return Utils_Date.defaultComparer(item.startTime, currentTime) <= 0;
            }
            else {
                return Utils_Date.defaultComparer(item.startTime, currentTime) <= 0 && Utils_Date.defaultComparer(item.endTime, currentTime) >= 0;
            }
        }).sort((b1: IBugBash, b2: IBugBash) => {
            return Utils_Date.defaultComparer(b1.startTime, b2.startTime);
        });
    }

    private _getUpcomingBugBashes(list: IBugBash[], currentTime: Date): IBugBash[] {
        return list.filter((item: IBugBash) => {
            return item.startTime && Utils_Date.defaultComparer(item.startTime, currentTime) > 0;
        }).sort((b1: IBugBash, b2: IBugBash) => {
            return Utils_Date.defaultComparer(b1.startTime, b2.startTime);
        });
    }
}