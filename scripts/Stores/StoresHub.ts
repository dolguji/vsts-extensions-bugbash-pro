import { StoreFactory } from "VSTS_Extension/Flux/Stores/BaseStore";
import { WorkItemFieldStore } from "VSTS_Extension/Flux/Stores/WorkItemFieldStore";
import { WorkItemTypeStore } from "VSTS_Extension/Flux/Stores/WorkItemTypeStore";
import { WorkItemTemplateStore } from "VSTS_Extension/Flux/Stores/WorkItemTemplateStore";
import { WorkItemTemplateItemStore } from "VSTS_Extension/Flux/Stores/WorkItemTemplateItemStore";
import { WorkItemStateItemStore } from "VSTS_Extension/Flux/Stores/WorkItemStateItemStore";
import { GitRepoStore } from "VSTS_Extension/Flux/Stores/GitRepoStore";
import { TeamStore } from "VSTS_Extension/Flux/Stores/TeamStore";
import { TeamFieldStore } from "VSTS_Extension/Flux/Stores/TeamFieldStore";

import { BugBashStore } from "./BugBashStore";
import { BugBashItemStore } from "./BugBashItemStore";
import { BugBashItemCommentStore } from "./BugBashItemCommentStore";
import { SettingsStore } from "./SettingsStore";
import { WorkItemStore } from "./WorkItemStore";

export module StoresHub {
    export var workItemStore: WorkItemStore = StoreFactory.getInstance<WorkItemStore>(WorkItemStore);
    export var settingsStore: SettingsStore = StoreFactory.getInstance<SettingsStore>(SettingsStore);
    export var bugBashStore: BugBashStore = StoreFactory.getInstance<BugBashStore>(BugBashStore);
    export var bugBashItemStore: BugBashItemStore = StoreFactory.getInstance<BugBashItemStore>(BugBashItemStore);
    export var bugBashItemCommentStore: BugBashItemCommentStore = StoreFactory.getInstance<BugBashItemCommentStore>(BugBashItemCommentStore);

    export var gitRepoStore: GitRepoStore = StoreFactory.getInstance<GitRepoStore>(GitRepoStore);
    export var teamStore: TeamStore = StoreFactory.getInstance<TeamStore>(TeamStore);
    export var teamFieldStore: TeamFieldStore = StoreFactory.getInstance<TeamFieldStore>(TeamFieldStore);
    export var workItemFieldStore: WorkItemFieldStore = StoreFactory.getInstance<WorkItemFieldStore>(WorkItemFieldStore);
    export var workItemTypeStore: WorkItemTypeStore = StoreFactory.getInstance<WorkItemTypeStore>(WorkItemTypeStore);
    export var workItemTemplateStore: WorkItemTemplateStore = StoreFactory.getInstance<WorkItemTemplateStore>(WorkItemTemplateStore);
    export var workItemTemplateItemStore: WorkItemTemplateItemStore = StoreFactory.getInstance<WorkItemTemplateItemStore>(WorkItemTemplateItemStore);
    export var workItemStateItemStore: WorkItemStateItemStore = StoreFactory.getInstance<WorkItemStateItemStore>(WorkItemStateItemStore);
}