import "../../css/BugBashResultsAnalytics.scss";

import * as React from "react";
import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { MessagePanel, MessageType } from "VSTS_Extension/Components/Common/MessagePanel";

import { Label } from "OfficeFabric/Label";
import { Bar, BarChart, XAxis, YAxis, Tooltip } from "recharts";

import { IBugBashItem } from "../Interfaces";
import { BugBashItemHelpers } from "../Helpers";
import { StoresHub } from "../Stores/StoresHub";

interface IBugBashResultsAnalyticsProps extends IBaseComponentProps {
    itemModels: IBugBashItem[];
}

const CustomizedAxisTick: React.StatelessComponent<any> =
    (props: any): JSX.Element => {
        const {x, y, stroke, payload} = props;
        const value = (payload.value.length > 9) ? payload.value.substr(0, 9) + "..." : payload.value;
        return (
            <g transform={`translate(${x-4},${y+2})`}>
                <text fill="#767676" style={{fontSize: "12px"}} width={100} textAnchor="end">{value}</text>
            </g>
        );
    };

export class BugBashResultsAnalytics extends BaseComponent<IBugBashResultsAnalyticsProps, IBaseComponentState> {
    protected initializeState() {
        this.state = {
            
        };
    }

    public render() {
        if (this.props.itemModels.length === 0) {
            return <MessagePanel messageType={MessageType.Info} message="No items created yet." />;
        }

        let teamCounts: IDictionaryStringTo<number> = {};
        let createdByCounts: IDictionaryStringTo<number> = {};
        let teamData = [];
        let createdByData = [];

        for (const model of this.props.itemModels) {
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
                    <Label className="header">Team</Label>
                    <BarChart layout={"vertical"} width={600} height={600} data={teamData} barSize={10}
                        margin={{top: 5, right: 30, left: 20, bottom: 5}}>
                        <XAxis type="number" allowDecimals={false} />
                        <YAxis type="category" dataKey="name" tick={<CustomizedAxisTick />} allowDecimals={false} />
                        <Tooltip />
                        <Bar dataKey="value" fill="#8884d8" />
                    </BarChart>
                </div>                
                <div className="chart-view">
                    <Label className="header">Created By</Label>
                    <BarChart layout={"vertical"} width={600} height={600} data={createdByData} barSize={10}
                        margin={{top: 5, right: 30, left: 20, bottom: 5}}>
                        <XAxis type="number" allowDecimals={false} />
                        <YAxis type="category" dataKey="name" tick={<CustomizedAxisTick />} allowDecimals={false} />
                        <Tooltip />
                        <Bar dataKey="value" fill="#8884d8" />
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