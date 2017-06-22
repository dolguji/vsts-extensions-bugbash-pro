import { BaseStore, StoreFactory } from "VSTS_Extension/Stores/BaseStore";
import { WorkItemFieldStore } from "VSTS_Extension/Stores/WorkItemFieldStore";
import { WorkItemTypeStore } from "VSTS_Extension/Stores/WorkItemTypeStore";
import { WorkItemTemplateStore } from "VSTS_Extension/Stores/WorkItemTemplateStore";
import { WorkItemTemplateItemStore } from "VSTS_Extension/Stores/WorkItemTemplateItemStore";

import { BugBashStore } from "../Stores/BugBashStore";
import { TeamStore } from "../Stores/TeamStore";
import { TeamFieldStore } from "../Stores/TeamFieldStore";

export module StoresHub {
    export var bugBashStore: BugBashStore = StoreFactory.getInstance<BugBashStore>(BugBashStore);
    export var teamStore: TeamStore = StoreFactory.getInstance<TeamStore>(TeamStore);
    export var teamFieldStore: TeamFieldStore = StoreFactory.getInstance<TeamFieldStore>(TeamFieldStore);
    export var workItemFieldStore: WorkItemFieldStore = StoreFactory.getInstance<WorkItemFieldStore>(WorkItemFieldStore);
    export var workItemTypeStore: WorkItemTypeStore = StoreFactory.getInstance<WorkItemTypeStore>(WorkItemTypeStore);
    export var workItemTemplateStore: WorkItemTemplateStore = StoreFactory.getInstance<WorkItemTemplateStore>(WorkItemTemplateStore);
    export var workItemTemplateItemStore: WorkItemTemplateItemStore = StoreFactory.getInstance<WorkItemTemplateItemStore>(WorkItemTemplateItemStore);
}