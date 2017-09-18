import { StoreFactory } from "MB/Flux/Stores/BaseStore";
import { WorkItemFieldStore } from "MB/Flux/Stores/WorkItemFieldStore";
import { WorkItemTypeStore } from "MB/Flux/Stores/WorkItemTypeStore";
import { WorkItemTemplateStore } from "MB/Flux/Stores/WorkItemTemplateStore";
import { WorkItemTemplateItemStore } from "MB/Flux/Stores/WorkItemTemplateItemStore";
import { WorkItemStateItemStore } from "MB/Flux/Stores/WorkItemStateItemStore";
import { GitRepoStore } from "MB/Flux/Stores/GitRepoStore";
import { TeamStore } from "MB/Flux/Stores/TeamStore";
import { TeamFieldStore } from "MB/Flux/Stores/TeamFieldStore";
import { WorkItemStore } from "MB/Flux/Stores/WorkItemStore";

import { BugBashStore } from "./BugBashStore";
import { BugBashErrorMessageStore } from "./BugBashErrorMessageStore";
import { BugBashItemStore } from "./BugBashItemStore";
import { BugBashItemCommentStore } from "./BugBashItemCommentStore";
import { BugBashSettingsStore } from "./BugBashSettingsStore";
import { UserSettingsStore } from "./UserSettingsStore";
import { LongTextStore } from "./LongTextStore";

export namespace StoresHub {
    export const bugBashSettingsStore: BugBashSettingsStore = StoreFactory.getInstance<BugBashSettingsStore>(BugBashSettingsStore);
    export const userSettingsStore: UserSettingsStore = StoreFactory.getInstance<UserSettingsStore>(UserSettingsStore);
    export const bugBashStore: BugBashStore = StoreFactory.getInstance<BugBashStore>(BugBashStore);
    export const bugBashItemStore: BugBashItemStore = StoreFactory.getInstance<BugBashItemStore>(BugBashItemStore);
    export const bugBashErrorMessageStore: BugBashErrorMessageStore = StoreFactory.getInstance<BugBashErrorMessageStore>(BugBashErrorMessageStore);
    export const bugBashItemCommentStore: BugBashItemCommentStore = StoreFactory.getInstance<BugBashItemCommentStore>(BugBashItemCommentStore);
    export const longTextStore: LongTextStore = StoreFactory.getInstance<LongTextStore>(LongTextStore);

    export const workItemStore: WorkItemStore = StoreFactory.getInstance<WorkItemStore>(WorkItemStore);
    export const gitRepoStore: GitRepoStore = StoreFactory.getInstance<GitRepoStore>(GitRepoStore);
    export const teamStore: TeamStore = StoreFactory.getInstance<TeamStore>(TeamStore);
    export const teamFieldStore: TeamFieldStore = StoreFactory.getInstance<TeamFieldStore>(TeamFieldStore);
    export const workItemFieldStore: WorkItemFieldStore = StoreFactory.getInstance<WorkItemFieldStore>(WorkItemFieldStore);
    export const workItemTypeStore: WorkItemTypeStore = StoreFactory.getInstance<WorkItemTypeStore>(WorkItemTypeStore);
    export const workItemTemplateStore: WorkItemTemplateStore = StoreFactory.getInstance<WorkItemTemplateStore>(WorkItemTemplateStore);
    export const workItemTemplateItemStore: WorkItemTemplateItemStore = StoreFactory.getInstance<WorkItemTemplateItemStore>(WorkItemTemplateItemStore);
    export const workItemStateItemStore: WorkItemStateItemStore = StoreFactory.getInstance<WorkItemStateItemStore>(WorkItemStateItemStore);
}