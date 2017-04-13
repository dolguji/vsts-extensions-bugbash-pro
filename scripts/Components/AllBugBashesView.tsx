import * as React from "react";
import { List } from "OfficeFabric/List";
import { autobind } from "OfficeFabric/Utilities";
import { Label } from "OfficeFabric/Label";
import { CommandBar } from "OfficeFabric/CommandBar";
import { IContextualMenuItem } from "OfficeFabric/components/ContextualMenu/ContextualMenu.Props";

import { HostNavigationService } from "VSS/SDK/Services/Navigation";
import Utils_Date = require("VSS/Utils/Date");

import { UrlActions, IBugBash, LoadingState } from "../Models";
import { HubView, IHubViewState } from "./HubView";
import { Loading } from "./Loading";
import { MessagePanel, MessageType } from "./MessagePanel";

interface IAllHubViewState extends IHubViewState {
    allItems: IBugBash[];
    pastItems: IBugBash[];
    currentItems: IBugBash[];
    upcomingItems: IBugBash[];
}

export class AllBugBashesView extends HubView<IAllHubViewState> {
    
    public render(): JSX.Element {
        return (
            <div className="all-view">
                <div className="all-view-contents">
                    <CommandBar className="all-view-menu-toolbar" items={this._getMenuItems()} />
                    {this._getContents()}
                </div>
            </div>
        );
    }

    private _getContents(): JSX.Element {
        if (this.state.loadingState === LoadingState.Loading) {
            return <Loading />;
        }
        else {
            if (this.state.allItems.length == 0) {
                return <MessagePanel message="No instance of bug bash exists in the context of current team." messageType={MessageType.Info} />
            }
            else {
                return (                    
                    <div className="instance-list-container">
                        <div className="instance-list-section">
                            <Label className="header">Past Bug Bashes ({this.state.pastItems.length})</Label>
                            <div className="instance-list-content">
                                {this.state.pastItems.length === 0 && <MessagePanel message="No past bug bashes." messageType={MessageType.Info} />}
                                {this.state.pastItems.length > 0 && <List items={this.state.pastItems} className="instance-list" onRenderCell={this._onRenderCell} />}
                            </div>
                        </div>

                        <div className="instance-list-section">
                            <Label className="header">Ongoing Bug Bashes ({this.state.currentItems.length})</Label>
                            <div className="instance-list-content">
                                {this.state.currentItems.length === 0 && <MessagePanel message="No ongoing bug bashes." messageType={MessageType.Info} />}
                                {this.state.currentItems.length > 0 && <List items={this.state.currentItems} className="instance-list" onRenderCell={this._onRenderCell} />}
                            </div>
                        </div>

                        <div className="instance-list-section">
                            <Label className="header">Upcoming Bug Bashes ({this.state.upcomingItems.length})</Label>
                            <div className="instance-list-content">
                                {this.state.upcomingItems.length === 0 && <MessagePanel message="No upcoming bug bashes." messageType={MessageType.Info} />}
                                {this.state.upcomingItems.length > 0 && <List items={this.state.upcomingItems} className="instance-list" onRenderCell={this._onRenderCell} />}
                            </div>
                        </div>
                    </div>
                );                
            }
        }
    }

    protected initialize(): void {
        this.props.context.actionsCreator.initializeAllBugBashes();
    }

    protected getStateFromStore(): IAllHubViewState {
        let allItems = this.props.context.stores.bugBashItemStore.getAll();
        let currentTime = new Date();
        return {
            allItems: allItems,
            pastItems: this._getPastBugBashes(allItems, currentTime),
            currentItems: this._getCurrentBugBashes(allItems, currentTime),
            upcomingItems: this._getUpcomingBugBashes(allItems, currentTime),
            loadingState: this.props.context.stores.bugBashItemStore.isLoaded() ? LoadingState.Loaded : LoadingState.Loading
        };
    }

    @autobind
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
                onClick: (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                    this.props.context.actionsCreator.refreshAllBugBashes();
                }
            }
         ];
    }

    @autobind
    private _onRenderCell(item: IBugBash, index?: number): React.ReactNode {
        return (
            <div className="instance-row">
                <div className="instance-title" onClick={() => this._onRowClick(item)}>{ item.title }</div>
                <div className="instance-info">
                    <div className="instance-info-cell-container"><div className="instance-info-cell">Start:</div><div className="instance-info-cell-info">{item.startTime ? Utils_Date.format(item.startTime, "dddd, MMMM dd, yyyy") : "N/A"}</div></div>
                    <div className="instance-info-cell-container"><div className="instance-info-cell">End:</div><div className="instance-info-cell-info">{item.endTime ? Utils_Date.format(item.endTime, "dddd, MMMM dd, yyyy") : "N/A"}</div></div>
                </div>
            </div>
        );
    }

    @autobind
    private async _onRowClick(item: IBugBash) {
        let navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
        navigationService.updateHistoryEntry(UrlActions.ACTION_VIEW, {id: item.id});
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