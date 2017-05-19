import "../../css/BugBashResultsAnalytics.scss";

import * as React from "react";
import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { IdentityView } from "VSTS_Extension/Components/WorkItemControls/IdentityView";

import { Label } from "OfficeFabric/Label";
import { MessageBar, MessageBarType } from "OfficeFabric/MessageBar";
import { Bar } from "recharts/lib/cartesian/Bar";
import { XAxis } from "recharts/lib/cartesian/XAxis";
import { YAxis } from "recharts/lib/cartesian/YAxis";
import { Text } from "recharts/lib/component/Text";
import { Tooltip } from "recharts/lib/component/Tooltip";
import { BarChart } from "recharts/lib/chart/BarChart";

import { WorkItem } from "TFS/WorkItemTracking/Contracts";

import { IBugBash, IBugBashItem } from "../Interfaces";
import { BugBashItemHelpers } from "../Helpers";

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

        let areaCounts: IDictionaryStringTo<number> = {};
        let createdByCounts: IDictionaryStringTo<number> = {};
        let areaData = [];
        let createdByData = [];

        for (const model of this.props.itemModels) {
            let areaPath = (BugBashItemHelpers.isAccepted(model) && this.props.workItemsMap[model.workItemId])
                ? this.props.workItemsMap[model.workItemId].fields["System.AreaPath"]
                : model.areaPath;

            let createdBy = model.createdBy;

            if (areaCounts[areaPath] == null) {
                areaCounts[areaPath] = 0;
            }
            if (createdByCounts[createdBy] == null) {
                createdByCounts[createdBy] = 0;
            }

            areaCounts[areaPath] = areaCounts[areaPath] + 1;
            createdByCounts[createdBy] = createdByCounts[createdBy] + 1;
        }

        for (const areaName in areaCounts) {
            areaData.push({ name: this._processAreaName(areaName), value: areaCounts[areaName]});
        }
        for (const createdBy in createdByCounts) {
            createdByData.push({ name: this._processIdentityName(createdBy), value: createdByCounts[createdBy]});
        }

        return <div className="bugbash-analytics">
                <div className="chart-view">
                    <Label className="header">Area Path</Label>
                    <BarChart layout={"vertical"} width={600} height={600} data={areaData} barSize={10}
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

    private _processAreaName(area: string): string {
        const i = area.indexOf("\\");
        return i > 0 ? area.substr(i + 1) : area;
    }

    private _processIdentityName(name: string): string {
        const i = name.indexOf(" <");
        return i > 0 ? name.substr(0, i) : name;
    }
}