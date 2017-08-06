import "../../css/BugBashCharts.scss";

import * as React from "react";
import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { BaseStore } from "VSTS_Extension/Flux/Stores/BaseStore";
import { Loading } from "VSTS_Extension/Components/Common/Loading";
import { TeamActions } from "VSTS_Extension/Flux/Actions/TeamActions";

import { MessageBar, MessageBarType } from "OfficeFabric/MessageBar";
import { Label } from "OfficeFabric/Label";
import { Bar, BarChart, XAxis, YAxis, Tooltip, CartesianGrid, ResponsiveContainer } from "recharts";

import { IBugBashItem, IBugBash } from "../Interfaces";
import { StoresHub } from "../Stores/StoresHub";
import { BugBashItemActions } from "../Actions/BugBashItemActions";
import { BugBashItemHelpers } from "../Helpers";
import { ChartsView } from "../Constants";

interface IBugBashChartsState extends IBaseComponentState {
    allBugBashItems: IBugBashItem[];
    pendingBugBashItems: IBugBashItem[];
    acceptedBugBashItems: IBugBashItem[];
    rejectedBugBashItems: IBugBashItem[];    
}

interface IBugBashChartsProps extends IBaseComponentProps {
    bugBash: IBugBash;
    view?: string;
}

interface INameValuePair {
    name: string;
    value: number;
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
        const data = props && props.payload && props.payload[0] && props.payload[0].payload;
        if (!data) {
            return null;
        }

        const keys = Object.keys(data).filter(k => k !== "name" && k !== "value");
        if (keys.length === 0) {
            return <div className="chart-tooltip"><span className="tooltip-key">{data["name"]}</span> : <span className="tooltip-value">{data["value"]}</span></div>;
        }
        else {
            return <div className="chart-tooltip">
                { keys.map(k => <div key={k}><span className="tooltip-key">{k}</span> : <span className="tooltip-value">{data[k]}</span></div>) }
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
            loading: true
        };
    }

    protected getStores(): BaseStore<any, any, any>[] {
        return [StoresHub.bugBashItemStore, StoresHub.teamStore];
    }

    protected getStoresState(): IBugBashChartsState {
        const bugBashItems = StoresHub.bugBashItemStore.getBugBashItems(this.props.bugBash.id);

        return {
            loading: StoresHub.bugBashItemStore.isLoading(this.props.bugBash.id) || StoresHub.teamStore.isLoading(),
            allBugBashItems: bugBashItems,
            pendingBugBashItems: bugBashItems ? bugBashItems.filter(b => !BugBashItemHelpers.isAccepted(b) && !b.rejected) : null,
            acceptedBugBashItems: bugBashItems ? bugBashItems.filter(b => BugBashItemHelpers.isAccepted(b)) : null,
            rejectedBugBashItems: bugBashItems ? bugBashItems.filter(b => !BugBashItemHelpers.isAccepted(b) && b.rejected) : null
        } as IBugBashChartsState;
    }

    public componentDidMount(): void {
        super.componentDidMount();
        TeamActions.initializeTeams();
        BugBashItemActions.initializeItems(this.props.bugBash.id);
    }  

    public render() {
        if (this.state.loading) {
            return <Loading />;
        }

        if (this.state.allBugBashItems.length === 0) {
            return <MessageBar messageBarType={MessageBarType.info} className="message-panel">
                No items created yet.
            </MessageBar>;            
        }

        let bugBashItems = this.state.allBugBashItems;
        if (this.props.view === ChartsView.AcceptedItemsOnly) {
            bugBashItems = this.state.acceptedBugBashItems;
        }
        else if (this.props.view === ChartsView.RejectedItemsOnly) {
            bugBashItems = this.state.rejectedBugBashItems;
        }
        else if (this.props.view === ChartsView.PendingItemsOnly) {
            bugBashItems = this.state.pendingBugBashItems;
        }

        let assignedToTeamCounts: IDictionaryStringTo<number> = {};
        let createdByCounts: IDictionaryStringTo<number> = {};
        let assignedToTeamData: INameValuePair[] = [];
        let createdByData: INameValuePair[] = [];

        for (const model of bugBashItems) {
            let teamId = model.teamId;

            let createdBy = model.createdBy;

            if (assignedToTeamCounts[teamId] == null) {
                assignedToTeamCounts[teamId] = 0;
            }
            if (createdByCounts[createdBy] == null) {
                createdByCounts[createdBy] = 0;
            }

            assignedToTeamCounts[teamId] = assignedToTeamCounts[teamId] + 1;
            createdByCounts[createdBy] = createdByCounts[createdBy] + 1;
        }

        for (const teamId in assignedToTeamCounts) {
            assignedToTeamData.push({ name: this._getTeamName(teamId), value: assignedToTeamCounts[teamId]});
        }
        for (const createdBy in createdByCounts) {
            createdByData.push({ name: this._getIdentityDisplayName(createdBy), value: createdByCounts[createdBy]});
        }

        const data = [
            {name: 'WITIQ', value:100, "Mohit Bagra": 10 , "Matt Manela": 100, "Karthik Bala": 50, "Mohit": 10 , "Matt": 100, "Karthik": 50, "M": 10 , "MM": 100, "K": 50},
            {name: 'MWIT', value:1000,"Mohit Bagra1": 100 , "Matt Manela1": 5, "Karthik Bala1": 20, "Mohit1": 10 , "Matt1": 100, "Karthik1": 50, "M1": 10 , "MM1": 100, "K1": 50},
            {name: 'WITPI', value:100,"Mohit Bagra2": 40 , "Matt Manela2": 100, "Karthik2": 50, "M2": 10 , "MM2": 100, "K2": 50},
            {name: 'WITX', value:50},
        ];
        
        assignedToTeamData.sort((a, b) => b.value - a.value);
        createdByData.sort((a, b) => b.value - a.value);
        data.sort((a, b) => b.value - a.value);

        return <div className="bugbash-analytics">
                <div className="chart-view">
                    <Label className="header">{`Assigned to team (${bugBashItems.length})`}</Label>
                    <ResponsiveContainer>
                    <BarChart layout={"vertical"} width={600} height={600} data={assignedToTeamData} barSize={10}
                        margin={{top: 5, right: 30, left: 20, bottom: 5}}>
                        <XAxis type="number" allowDecimals={false} />
                        <YAxis type="category" dataKey="name" tick={<CustomAxisTick />} allowDecimals={false} />
                        <CartesianGrid strokeDasharray="3 3"/>
                        <Tooltip isAnimationActive={false} />
                        <Bar isAnimationActive={false} dataKey="value" fill="#8884d8" />
                    </BarChart>
                    </ResponsiveContainer>
                </div>                
                <div className="chart-view">
                    <Label className="header">{`Created By (${bugBashItems.length})`}</Label>
                    <BarChart layout={"vertical"} width={600} height={600} data={data} barSize={10}
                        margin={{top: 5, right: 30, left: 20, bottom: 5}}>
                        <XAxis type="number" allowDecimals={false} />
                        <YAxis type="category" dataKey="name" tick={<CustomAxisTick />} allowDecimals={false} />
                        <CartesianGrid strokeDasharray="3 3"/>
                        <Tooltip isAnimationActive={false} content={<CustomTooltip/>}/>
                        <Bar isAnimationActive={false} dataKey="value" fill="#8884d8" />
                    </BarChart>
                </div>
            </div>
    }

    private _getTeamName(teamId: string): string {
        const team = StoresHub.teamStore.getItem(teamId);
        return team ? team.name : teamId;
    }

    private _getIdentityDisplayName(name: string): string {
        const i = name.indexOf(" <");
        return i > 0 ? name.substr(0, i) : name;
    }
}