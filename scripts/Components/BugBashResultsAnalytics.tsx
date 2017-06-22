import "../../css/BugBashResultsAnalytics.scss";

import * as React from "react";
import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";

import { Label } from "OfficeFabric/Label";
import { MessageBar, MessageBarType } from "OfficeFabric/MessageBar";
import { Bar, BarChart, XAxis, YAxis, Text, Tooltip } from "recharts";

import { WorkItem } from "TFS/WorkItemTracking/Contracts";

import { IBugBashItem } from "../Interfaces";
import { BugBashItemHelpers } from "../Helpers";
import { StoresHub } from "../Stores/StoresHub";

interface IBugBashResultsAnalyticsProps extends IBaseComponentProps {
    bugBashId: string;
    itemModels: IBugBashItem[];
    workItemsMap: IDictionaryNumberTo<WorkItem>;
}

const CustomizedAxisTick: React.StatelessComponent<any> =
    (props: any): JSX.Element => {
        const {x, y, stroke, payload} = props;		
        return (
            <g transform={`translate(${x},${y})`}>
                <Text angle={-35} width={100} textAnchor="end" verticalAnchor="middle">{payload.value}</Text>
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
            return <MessageBar messageBarType={MessageBarType.info}>No items created yet.</MessageBar>;
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