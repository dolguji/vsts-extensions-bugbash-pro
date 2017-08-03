import "../../css/BugBashCharts.scss";

import * as React from "react";
import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { BaseStore } from "VSTS_Extension/Flux/Stores/BaseStore";
import { Loading } from "VSTS_Extension/Components/Common/Loading";
import { TeamActions } from "VSTS_Extension/Flux/Actions/TeamActions";

import { MessageBar, MessageBarType } from "OfficeFabric/MessageBar";
import { Label } from "OfficeFabric/Label";
import { Bar, BarChart, XAxis, YAxis, Tooltip } from "recharts";

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

const CustomizedAxisTick: React.StatelessComponent<any> =
    (props: any): JSX.Element => {
        const {x, y, payload} = props;
        const value = (payload.value.length > 9) ? payload.value.substr(0, 9) + "..." : payload.value;
        return (
            <g transform={`translate(${x-4},${y+2})`}>
                <text fill="#767676" style={{fontSize: "12px"}} width={100} textAnchor="end">{value}</text>
            </g>
        );
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

        let teamCounts: IDictionaryStringTo<number> = {};
        let createdByCounts: IDictionaryStringTo<number> = {};
        let teamData = [];
        let createdByData = [];

        for (const model of bugBashItems) {
            let teamId = model.teamId;

            let createdBy = model.createdBy;

            if (teamCounts[teamId] == null) {
                teamCounts[teamId] = 0;
            }
            if (createdByCounts[createdBy] == null) {
                createdByCounts[createdBy] = 0;
            }

            teamCounts[teamId] = teamCounts[teamId] + 1;
            createdByCounts[createdBy] = createdByCounts[createdBy] + 1;
        }

        for (const teamId in teamCounts) {
            teamData.push({ name: this._processTeamName(teamId), value: teamCounts[teamId]});
        }
        for (const createdBy in createdByCounts) {
            createdByData.push({ name: this._processIdentityName(createdBy), value: createdByCounts[createdBy]});
        }

        return <div className="bugbash-analytics">
                <div className="chart-view">
                    <Label className="header">{`Assigned to team (${bugBashItems.length})`}</Label>
                    <BarChart layout={"vertical"} width={600} height={600} data={teamData} barSize={10}
                        margin={{top: 5, right: 30, left: 20, bottom: 5}}>
                        <XAxis type="number" allowDecimals={false} />
                        <YAxis type="category" dataKey="name" tick={<CustomizedAxisTick />} allowDecimals={false} />
                        <Tooltip />
                        <Bar isAnimationActive={false} dataKey="value" fill="#8884d8" />
                    </BarChart>
                </div>                
                <div className="chart-view">
                    <Label className="header">{`Created By (${bugBashItems.length})`}</Label>
                    <BarChart layout={"vertical"} width={600} height={600} data={createdByData} barSize={10}
                        margin={{top: 5, right: 30, left: 20, bottom: 5}}>
                        <XAxis type="number" allowDecimals={false} />
                        <YAxis type="category" dataKey="name" tick={<CustomizedAxisTick />} allowDecimals={false} />
                        <Tooltip />
                        <Bar isAnimationActive={false} dataKey="value" fill="#8884d8" />
                    </BarChart>
                </div>
            </div>
    }

    private _processTeamName(teamId: string): string {
        const team = StoresHub.teamStore.getItem(teamId);
        return team ? team.name : teamId;
    }

    private _processIdentityName(name: string): string {
        const i = name.indexOf(" <");
        return i > 0 ? name.substr(0, i) : name;
    }
}