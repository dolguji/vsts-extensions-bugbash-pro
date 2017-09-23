import "../../css/AllBugBashes.scss";

import * as React from "react";

import { autobind } from "OfficeFabric/Utilities";
import { SelectionMode } from "OfficeFabric/utilities/selection";
import { TooltipHost, TooltipDelay, DirectionalHint, TooltipOverflowMode } from "OfficeFabric/Tooltip";
import { Label } from "OfficeFabric/Label";
import { IContextualMenuItem } from "OfficeFabric/components/ContextualMenu/ContextualMenu.Props";
import { Panel, PanelType } from "OfficeFabric/Panel";
import { MessageBar, MessageBarType } from "OfficeFabric/MessageBar";

import { Loading } from "MB/Components/Loading";
import { Grid, GridColumn } from "MB/Components/Grid";
import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "MB/Components/BaseComponent";
import { BaseStore } from "MB/Flux/Stores/BaseStore";
import { Hub } from "MB/Components/Hub";
import { getAsyncLoadedComponent } from "MB/Components/AsyncLoadedComponent";
import { DateUtils } from "MB/Utils/Date";
import { StringUtils } from "MB/Utils/String";

import { confirmAction, getBugBashUrl, navigate } from "../Helpers";
import { UrlActions, ErrorKeys, BugBashFieldNames } from "../Constants";
import { StoresHub } from "../Stores/StoresHub";
import { BugBashActions } from "../Actions/BugBashActions";
import { BugBashErrorMessageActions } from "../Actions/BugBashErrorMessageActions";
import * as SettingsPanel_Async from "./SettingsPanel";
import { BugBash } from "../ViewModels/BugBash";

interface IAllBugBashesViewState extends IBaseComponentState {
    pastBugBashes: BugBash[];
    currentBugBashes: BugBash[];
    upcomingBugBashes: BugBash[];
    settingsPanelOpen: boolean;
    selectedPivot?: string;
    error?: string;    
}

const AsyncSettingsPanel = getAsyncLoadedComponent(
    ["scripts/SettingsPanel"],
    (m: typeof SettingsPanel_Async) => m.SettingsPanel,
    () => <Loading />);

class BugBashGrid extends Grid<BugBash> {};

export class AllBugBashesView extends BaseComponent<IBaseComponentProps, IAllBugBashesViewState> {
    protected getStores(): BaseStore<any, any, any>[] {
        return [StoresHub.bugBashStore, StoresHub.bugBashErrorMessageStore];
    }

    protected initializeState() {
        this.state = {
            pastBugBashes: [],
            currentBugBashes: [],
            upcomingBugBashes: [],
            loading: true,
            selectedPivot: "ongoing",
            settingsPanelOpen: false,
            error: StoresHub.bugBashErrorMessageStore.getItem(ErrorKeys.DirectoryPageError)
        };
    }

    public componentDidMount() {
        super.componentDidMount();
        BugBashActions.initializeAllBugBashes();     
    }

    public componentWillUnmount() {
        super.componentWillUnmount();
        this._dismissErrorMessage();
    }

    protected getStoresState(): IAllBugBashesViewState {        
        const allBugBashes = (StoresHub.bugBashStore.getAll() || [])
            .filter(bugBash => StringUtils.equals(VSS.getWebContext().project.id, bugBash.projectId, true));
        const currentTime = new Date();
        
        return {
            pastBugBashes: this._getPastBugBashes(allBugBashes, currentTime),
            currentBugBashes: this._getCurrentBugBashes(allBugBashes, currentTime),
            upcomingBugBashes: this._getUpcomingBugBashes(allBugBashes, currentTime),
            loading: StoresHub.bugBashStore.isLoading(),
            error: StoresHub.bugBashErrorMessageStore.getItem(ErrorKeys.DirectoryPageError)
        } as IAllBugBashesViewState;
    }

    public render(): JSX.Element {
        return (
            <div className="all-view">                
                { this.state.error && 
                    <MessageBar 
                        className="bugbash-error"
                        messageBarType={MessageBarType.error} 
                        onDismiss={this._dismissErrorMessage}>
                        {this.state.error}
                    </MessageBar>
                }
                <Hub 
                    className={this.state.error ? "with-error" : ""}
                    title="Bug Bashes"          
                    pivotProps={{
                        initialSelectedKey: this.state.selectedPivot,
                        onPivotClick: (selectedPivotKey: string) => {
                            this.updateState({selectedPivot: selectedPivotKey} as IAllBugBashesViewState);
                        },
                        onRenderPivotContent: (key: string) => this._getContents(key),
                        pivots: [
                            {
                                key: "ongoing",
                                text: "Ongoing",
                                itemCount: this.state.currentBugBashes ? this.state.currentBugBashes.length : null,
                                commands: this._getCommandBarItems()
                            },
                            {
                                key: "upcoming",
                                text: "Upcoming",
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
                        onDismiss={this._dismissSettingsPanel}>

                        <AsyncSettingsPanel />
                    </Panel>
                }
            </div>
        );
    }

    @autobind
    private _dismissSettingsPanel() {
        this.updateState({settingsPanelOpen: false} as IAllBugBashesViewState);
    }

    @autobind
    private _dismissErrorMessage() {
        BugBashErrorMessageActions.dismissErrorMessage(ErrorKeys.DirectoryPageError);
    };

    private _getContents(key: string): JSX.Element {
        if (this.state.loading) {
            return <Loading />;
        }
            
        let bugBashes: BugBash[];
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

        if (bugBashes.length === 0) {
            return <MessageBar messageBarType={MessageBarType.info} className="message-panel">
                {missingItemsMsg}
            </MessageBar>
        }

        return <BugBashGrid
            className={"instance-list"}
            items={bugBashes}
            columns={this._getGridColumns()}
            selectionMode={SelectionMode.none}
            getContextMenuItems={this._getGridContextMenuItems}
        />;                    
    }

    private _getGridColumns(): GridColumn<BugBash>[] {        
        return [
            {
                key: "title",
                name: "Title",
                minWidth: 300,
                maxWidth: Infinity,
                onRenderCell: (bugBash: BugBash) => {
                    const title = bugBash.getFieldValue<string>(BugBashFieldNames.Title, true);
                    return <TooltipHost 
                        content={title}
                        delay={TooltipDelay.medium}
                        overflowMode={TooltipOverflowMode.Parent}
                        directionalHint={DirectionalHint.bottomLeftEdge}>
                        <Label className="bugbash-grid-cell">
                            <a href={getBugBashUrl(bugBash.id, UrlActions.ACTION_RESULTS)} onClick={(e: React.MouseEvent<HTMLElement>) => this._onRowClick(e, bugBash)}>{ title }</a>
                        </Label>
                    </TooltipHost>;
                }
            },
            {
                key: "startDate",
                name: "Start Date",
                minWidth: 400,
                maxWidth: 600,
                onRenderCell: (bugBash: BugBash) => {
                    const startTime = bugBash.getFieldValue<Date>(BugBashFieldNames.StartTime, true);
                    const label = startTime ? DateUtils.format(startTime, "dddd, mmmm dd, yyyy") : "N/A";
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
                onRenderCell: (bugBash: BugBash) => {
                    const endTime = bugBash.getFieldValue<Date>(BugBashFieldNames.EndTime, true);
                    const label = endTime ? DateUtils.format(endTime, "dddd, mmmm dd, yyyy") : "N/A";

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

    @autobind
    private _getGridContextMenuItems(selectedItems: BugBash[]): IContextualMenuItem[] {
        if (selectedItems.length !== 1) {
            return [];
        }

        const bugBash = selectedItems[0];
        return [
            {
                key: "open", name: "View results", iconProps: {iconName: "ShowResults"}, 
                onClick: () => {                    
                    navigate(UrlActions.ACTION_RESULTS, {id: bugBash.id});
                }
            },
            {
                key: "edit", name: "Edit", iconProps: {iconName: "Edit"}, 
                onClick: () => {                    
                    navigate(UrlActions.ACTION_EDIT, {id: bugBash.id});
                }
            },
            {
                key: "delete", name: "Delete", iconProps: {iconName: "Cancel", style: { color: "#da0a00", fontWeight: "bold" }}, 
                onClick: async () => {                    
                    const confirm = await confirmAction(true, "Are you sure you want to delete this bug bash instance?");
                    if (confirm) {
                        bugBash.delete();
                    }   
                }
            },
        ];
    }

    private _getCommandBarItems(): IContextualMenuItem[] {
         return [
            {
                key: "new", name: "New Bug Bash", iconProps: {iconName: "Add"},
                onClick: () => {
                    navigate(UrlActions.ACTION_EDIT);
                }
            },            
            {
                key: "refresh", name: "Refresh", iconProps: {iconName: "Refresh"},
                onClick: () => {
                    BugBashActions.refreshAllBugBashes();
                }
            },
            {
                key: "settings", name: "Settings", iconProps: {iconName: "Settings"},
                onClick: () => {
                    this.updateState({settingsPanelOpen: !(this.state.settingsPanelOpen)} as IAllBugBashesViewState);
                }
            }
         ];
    }

    private _onRowClick(e: React.MouseEvent<HTMLElement>, bugBash: BugBash) {
        if (!e.ctrlKey) {
            e.preventDefault();
            navigate(UrlActions.ACTION_RESULTS, {id: bugBash.id});
        }      
    }

    private _getPastBugBashes(list: BugBash[], currentTime: Date): BugBash[] {
        return list.filter((bugBash: BugBash) => {
            const endTime = bugBash.getFieldValue<Date>(BugBashFieldNames.EndTime, true);
            return endTime && DateUtils.defaultComparer(endTime, currentTime) < 0;
        }).sort((b1: BugBash, b2: BugBash) => {
            const endTime1 = b1.getFieldValue<Date>(BugBashFieldNames.EndTime, true);
            const endTime2 = b2.getFieldValue<Date>(BugBashFieldNames.EndTime, true);
            return DateUtils.defaultComparer(endTime1, endTime2);
        });
    }

    private _getCurrentBugBashes(list: BugBash[], currentTime: Date): BugBash[] {        
        return list.filter((bugBash: BugBash) => {
            const startTime = bugBash.getFieldValue<Date>(BugBashFieldNames.StartTime, true);
            const endTime = bugBash.getFieldValue<Date>(BugBashFieldNames.EndTime, true);

            if (!startTime && !endTime) {
                return true;
            }
            else if(!startTime && endTime) {
                return DateUtils.defaultComparer(endTime, currentTime) >= 0;
            }
            else if (startTime && !endTime) {
                return DateUtils.defaultComparer(startTime, currentTime) <= 0;
            }
            else {
                return DateUtils.defaultComparer(startTime, currentTime) <= 0 && DateUtils.defaultComparer(endTime, currentTime) >= 0;
            }
        }).sort((b1: BugBash, b2: BugBash) => {
            const startTime1 = b1.getFieldValue<Date>(BugBashFieldNames.StartTime, true);
            const startTime2 = b2.getFieldValue<Date>(BugBashFieldNames.StartTime, true);
            return DateUtils.defaultComparer(startTime1, startTime2);
        });
    }

    private _getUpcomingBugBashes(list: BugBash[], currentTime: Date): BugBash[] {
        return list.filter((bugBash: BugBash) => {
            const startTime = bugBash.getFieldValue<Date>(BugBashFieldNames.StartTime, true);
            return startTime && DateUtils.defaultComparer(startTime, currentTime) > 0;
        }).sort((b1: BugBash, b2: BugBash) => {
            const startTime1 = b1.getFieldValue<Date>(BugBashFieldNames.StartTime, true);
            const startTime2 = b2.getFieldValue<Date>(BugBashFieldNames.StartTime, true);
            return DateUtils.defaultComparer(startTime1, startTime2);
        });
    }
}