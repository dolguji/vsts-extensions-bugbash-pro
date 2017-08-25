import { StoreFactory } from "VSTS_Extension/Flux/Stores/BaseStore";
import { WorkItemFieldStore } from "VSTS_Extension/Flux/Stores/WorkItemFieldStore";
import { WorkItemTypeStore } from "VSTS_Extension/Flux/Stores/WorkItemTypeStore";
import { WorkItemTemplateStore } from "VSTS_Extension/Flux/Stores/WorkItemTemplateStore";
import { WorkItemTemplateItemStore } from "VSTS_Extension/Flux/Stores/WorkItemTemplateItemStore";
import { WorkItemStateItemStore } from "VSTS_Extension/Flux/Stores/WorkItemStateItemStore";
import { GitRepoStore } from "VSTS_Extension/Flux/Stores/GitRepoStore";
import { TeamStore } from "VSTS_Extension/Flux/Stores/TeamStore";
import { TeamFieldStore } from "VSTS_Extension/Flux/Stores/TeamFieldStore";
import { WorkItemStore } from "VSTS_Extension/Flux/Stores/WorkItemStore";

import { BugBashStore } from "./BugBashStore";
import { BugBashErrorMessageStore } from "./BugBashErrorMessageStore";
import { BugBashItemStore } from "./BugBashItemStore";
import { BugBashItemCommentStore } from "./BugBashItemCommentStore";
import { BugBashSettingsStore } from "./BugBashSettingsStore";
import { UserSettingsStore } from "./UserSettingsStore";

export namespace StoresHub {
    export const workItemStore: WorkItemStore = StoreFactory.getInstance<WorkItemStore>(WorkItemStore);
    export const bugBashSettingsStore: BugBashSettingsStore = StoreFactory.getInstance<BugBashSettingsStore>(BugBashSettingsStore);
    export const userSettingsStore: UserSettingsStore = StoreFactory.getInstance<UserSettingsStore>(UserSettingsStore);
    export const bugBashStore: BugBashStore = StoreFactory.getInstance<BugBashStore>(BugBashStore);
    export const bugBashItemStore: BugBashItemStore = StoreFactory.getInstance<BugBashItemStore>(BugBashItemStore);
    export const bugBashErrorMessageStore: BugBashErrorMessageStore = StoreFactory.getInstance<BugBashErrorMessageStore>(BugBashErrorMessageStore);
    export const bugBashItemCommentStore: BugBashItemCommentStore = StoreFactory.getInstance<BugBashItemCommentStore>(BugBashItemCommentStore);

    export const gitRepoStore: GitRepoStore = StoreFactory.getInstance<GitRepoStore>(GitRepoStore);
    export const teamStore: TeamStore = StoreFactory.getInstance<TeamStore>(TeamStore);
    export const teamFieldStore: TeamFieldStore = StoreFactory.getInstance<TeamFieldStore>(TeamFieldStore);
    export const workItemFieldStore: WorkItemFieldStore = StoreFactory.getInstance<WorkItemFieldStore>(WorkItemFieldStore);
    export const workItemTypeStore: WorkItemTypeStore = StoreFactory.getInstance<WorkItemTypeStore>(WorkItemTypeStore);
    export const workItemTemplateStore: WorkItemTemplateStore = StoreFactory.getInstance<WorkItemTemplateStore>(WorkItemTemplateStore);
    export const workItemTemplateItemStore: WorkItemTemplateItemStore = StoreFactory.getInstance<WorkItemTemplateItemStore>(WorkItemTemplateItemStore);
    export const workItemStateItemStore: WorkItemStateItemStore = StoreFactory.getInstance<WorkItemStateItemStore>(WorkItemStateItemStore);
}