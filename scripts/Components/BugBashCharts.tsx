import "../../css/BugBashCharts.scss";

import * as React from "react";
import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "MB/Components/BaseComponent";
import { BaseStore } from "MB/Flux/Stores/BaseStore";
import { Loading } from "MB/Components/Loading";
import { parseUniquefiedIdentityName } from "MB/Components/IdentityView";
import { TeamActions } from "MB/Flux/Actions/TeamActions";

import { PrimaryButton } from "OfficeFabric/Button";
import { MessageBar, MessageBarType } from "OfficeFabric/MessageBar";
import { Label } from "OfficeFabric/Label";
import { Checkbox } from "OfficeFabric/Checkbox";
import { Bar, BarChart, XAxis, YAxis, Tooltip, CartesianGrid, ResponsiveContainer } from "recharts";

import { INameValuePair } from "../Interfaces";
import { StoresHub } from "../Stores/StoresHub";
import { BugBashItemActions } from "../Actions/BugBashItemActions";
import { ChartsView, BugBashFieldNames, BugBashItemFieldNames } from "../Constants";
import { SettingsActions } from "../Actions/SettingsActions";
import { BugBash } from "../ViewModels/BugBash";
import * as ExcelExporter_Async from "../ExcelExporter";
import { BugBashItem } from "../ViewModels/BugBashItem";

interface IBugBashChartsState extends IBaseComponentState {
    allBugBashItems: BugBashItem[];
    pendingBugBashItems: BugBashItem[];
    acceptedBugBashItems: BugBashItem[];
    rejectedBugBashItems: BugBashItem[];   
    groupedByTeam?: boolean; 
}

interface IBugBashChartsProps extends IBaseComponentProps {
    bugBash: BugBash;
    view?: string;
}

const CustomAxisTick: React.StatelessComponent<any> =
    (props: any): JSX.Element => {
        const {x, y, payload} = props;
        const value = (payload.value.length > 9) ? payload.value.substr(0, 9) + "..." : payload.value;
        return (
            <g transform={`translate(${x-4},${y+2})`}>
                <text fill="#767676" style={{fontSize: "12px"}} width={100} textAnchor="end">{value}</text>
            </g>
        );
    };

const CustomTooltip: React.StatelessComponent<any> =
    (props: any): JSX.Element => {
        const data: INameValuePair = props && props.payload && props.payload[0] && props.payload[0].payload;
        if (!data) {
            return null;
        }
        
        if (!data.members || data.members.length === 0) {
            return <div className="chart-tooltip"><span className="tooltip-key">{data["name"]}</span> : <span className="tooltip-value">{data["value"]}</span></div>;
        }
        else {
            return <div className="chart-tooltip">
                <div className="team-name"><span className="tooltip-key">{data["name"]}</span> : <span className="tooltip-value">{data["value"]}</span></div>
                { data.members.map((member: INameValuePair) => {
                    return <div key={member.name}><span className="tooltip-key">{member.name}</span> : <span className="tooltip-value">{member.value}</span></div>
                })}
            </div>;
        }        
    };

export class BugBashCharts extends BaseComponent<IBugBashChartsProps, IBugBashChartsState> {
    protected initializeState() {
        this.state = {
            allBugBashItems: null,
            pendingBugBashItems: null,
            rejectedBugBashItems: null,
            acceptedBugBashItems: null,
            loading: true,
            groupedByTeam: true
        };
    }

    protected getStores(): BaseStore<any, any, any>[] {
        return [StoresHub.bugBashItemStore, StoresHub.teamStore, StoresHub.userSettingsStore];
    }

    protected getStoresState(): IBugBashChartsState {
        const bugBashItems = StoresHub.bugBashItemStore.getBugBashItems(this.props.bugBash.id);

        return {
            loading: StoresHub.bugBashItemStore.isLoading(this.props.bugBash.id) || StoresHub.teamStore.isLoading() || StoresHub.userSettingsStore.isLoading(),
            allBugBashItems: bugBashItems,
            pendingBugBashItems: bugBashItems ? bugBashItems.filter(b => !b.isAccepted && !b.getFieldValue<boolean>(BugBashItemFieldNames.Rejected, true)) : null,
            acceptedBugBashItems: bugBashItems ? bugBashItems.filter(b => b.isAccepted) : null,
            rejectedBugBashItems: bugBashItems ? bugBashItems.filter(b => !b.isAccepted && b.getFieldValue<boolean>(BugBashItemFieldNames.Rejected, true)) : null
        } as IBugBashChartsState;
    }

    public componentDidMount() {
        super.componentDidMount();
        TeamActions.initializeTeams();
        BugBashItemActions.initializeItems(this.props.bugBash.id);
        SettingsActions.initializeUserSettings();
    }  

    public render() {
        if (this.state.loading) {
            return <Loading />;
        }        

        if (this.state.allBugBashItems.length === 0) {
            return <MessageBar messageBarType={MessageBarType.info} className="message-panel">
                No items created yet
            </MessageBar>;            
        }

        let bugBashItems: BugBashItem[] = this.state.allBugBashItems;
        let emptyMessage = "No items created yet";
        if (this.props.view === ChartsView.AcceptedItemsOnly) {
            bugBashItems = this.state.acceptedBugBashItems;
            emptyMessage = "No item has been accepted yet.";
        }
        else if (this.props.view === ChartsView.RejectedItemsOnly) {
            bugBashItems = this.state.rejectedBugBashItems;
            emptyMessage = "No item has been rejected yet.";
        }
        else if (this.props.view === ChartsView.PendingItemsOnly) {
            bugBashItems = this.state.pendingBugBashItems;
            emptyMessage = "No pending items left.";
        }
        
        if (bugBashItems.length === 0) {
            return <MessageBar messageBarType={MessageBarType.info} className="message-panel">
                {emptyMessage}
            </MessageBar>;            
        }

        let assignedToTeamCounts: IDictionaryStringTo<number> = {};
        let createdByCounts: IDictionaryStringTo<{count: number, members: IDictionaryStringTo<number>}> = {};
        let assignedToTeamData: INameValuePair[] = [];
        let createdByData: INameValuePair[] = [];

        for (const model of bugBashItems) {
            const teamId = model.getFieldValue<string>(BugBashItemFieldNames.TeamId, true);
            const createdBy = model.getFieldValue<string>(BugBashItemFieldNames.CreatedBy, true);
            const createdByUser = parseUniquefiedIdentityName(createdBy);
            const userSetting = StoresHub.userSettingsStore.getItem(createdByUser.uniqueName);
            const associatedTeamId = userSetting ? userSetting.associatedTeam : "";
            const associatedTeam = associatedTeamId ? StoresHub.teamStore.getItem(associatedTeamId) : null;

            assignedToTeamCounts[teamId] = (assignedToTeamCounts[teamId] || 0) + 1;

            if (associatedTeam && this.state.groupedByTeam) {
                if (createdByCounts[associatedTeam.name] == null) {
                    createdByCounts[associatedTeam.name] = {
                        count: 0,
                        members: {}
                    };
                }
                createdByCounts[associatedTeam.name].count = createdByCounts[associatedTeam.name].count + 1;
                createdByCounts[associatedTeam.name].members[createdBy] = (createdByCounts[associatedTeam.name].members[createdBy] || 0) + 1;
            }
            else {
                if (createdByCounts[createdBy] == null) {
                    createdByCounts[createdBy] = {
                        count: 0,
                        members: {}
                    };
                }
                createdByCounts[createdBy].count = createdByCounts[createdBy].count + 1;
            }
        }

        for (const teamId in assignedToTeamCounts) {
            assignedToTeamData.push({ name: this._getTeamName(teamId), value: assignedToTeamCounts[teamId]});
        }

        for (const createdBy in createdByCounts) {
            let membersMap = createdByCounts[createdBy].members;
            let membersArr: INameValuePair[] = $.map(membersMap, (count: number, key: string) => {
                return {name: parseUniquefiedIdentityName(key).displayName, value: count}
            })
            membersArr.sort((a, b) => b.value - a.value);

            createdByData.push({ name: parseUniquefiedIdentityName(createdBy).displayName, value: createdByCounts[createdBy].count, members: membersArr});
        }

        assignedToTeamData.sort((a, b) => b.value - a.value);
        createdByData.sort((a, b) => b.value - a.value);

        return <div className="bugbash-charts">
                <div className="chart-view-container">
                    <div className="header-container">
                        <Label className="header">{`Assigned to team (${bugBashItems.length})`}</Label>
                         <PrimaryButton className="export-excel" onClick={() => {
                            requirejs(["scripts/ExcelExporter"], (ExcelExporter: typeof ExcelExporter_Async) => {
                                new ExcelExporter.ExcelExporter(assignedToTeamData).export(`${this.props.bugBash.getFieldValue<string>(BugBashFieldNames.Title, true)} - Assigned To Team.xlsx`);          
                            })
                        }}>
                            Export
                        </PrimaryButton> 
                    </div>
                    <div className="chart-view">
                        <ResponsiveContainer>
                            <BarChart layout={"vertical"} width={600} height={600} data={assignedToTeamData} barSize={5}
                                margin={{top: 5, right: 30, left: 20, bottom: 5}}>
                                <XAxis type="number" allowDecimals={false} />
                                <YAxis type="category" dataKey="name" tick={<CustomAxisTick />} allowDecimals={false} />
                                <CartesianGrid strokeDasharray="3 3"/>
                                <Tooltip isAnimationActive={false} />
                                <Bar isAnimationActive={false} dataKey="value" fill="#8884d8" />
                            </BarChart>
                        </ResponsiveContainer>
                    </div>
                </div>                
                <div className="chart-view-container">
                    <div className="header-container">
                        <Label className="header">{`Created By (${bugBashItems.length})`}</Label>
                        <Checkbox 
                            label="Group by team" 
                            checked={this.state.groupedByTeam}
                            className="group-by-checkbox"
                            onChange={() => {
                                this.updateState({groupedByTeam: !this.state.groupedByTeam} as IBugBashChartsState);
                            }}
                        />
                         <PrimaryButton className="export-excel" onClick={() => {
                            requirejs(["scripts/ExcelExporter"], (ExcelExporter: typeof ExcelExporter_Async) => {
                                new ExcelExporter.ExcelExporter(createdByData).export(`${this.props.bugBash.getFieldValue<string>(BugBashFieldNames.Title, true)} - Created By.xlsx`);          
                            })
                        }}>
                            Export
                        </PrimaryButton> 
                    </div>
                    <div className="chart-view">
                        <ResponsiveContainer>
                            <BarChart layout={"vertical"} width={600} height={600} data={createdByData} barSize={5}
                                margin={{top: 5, right: 30, left: 20, bottom: 5}}>
                                <XAxis type="number" allowDecimals={false} />
                                <YAxis type="category" dataKey="name" tick={<CustomAxisTick />} allowDecimals={false} />
                                <CartesianGrid strokeDasharray="3 3"/>
                                <Tooltip isAnimationActive={false} content={<CustomTooltip/>}/>
                                <Bar isAnimationActive={false} dataKey="value" fill="#8884d8" />
                            </BarChart>
                        </ResponsiveContainer>
                    </div>
                </div>
            </div>
    }

    private _getTeamName(teamId: string): string {
        const team = StoresHub.teamStore.getItem(teamId);
        return team ? team.name : teamId;
    }    
}
