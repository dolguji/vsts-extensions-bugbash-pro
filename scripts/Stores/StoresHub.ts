import { BaseStore, StoreFactory } from "VSTS_Extension/Stores/BaseStore";
import { WorkItemFieldStore } from "VSTS_Extension/Stores/WorkItemFieldStore";
import { WorkItemTypeStore } from "VSTS_Extension/Stores/WorkItemTypeStore";
import { WorkItemTemplateStore } from "VSTS_Extension/Stores/WorkItemTemplateStore";
import { WorkItemTemplateItemStore } from "VSTS_Extension/Stores/WorkItemTemplateItemStore";

import { BugBashStore } from "../Stores/BugBashStore";
import { BugBashItemStore } from "../Stores/BugBashItemStore";
import { BugBashItemCommentStore } from "../Stores/BugBashItemCommentStore";

export module StoresHub {
    export var bugBashStore: BugBashStore = StoreFactory.getInstance<BugBashStore>(BugBashStore);
    export var bugBashItemStore: BugBashItemStore = StoreFactory.getInstance<BugBashItemStore>(BugBashItemStore);
    export var bugBashItemCommentStore: BugBashItemCommentStore = StoreFactory.getInstance<BugBashItemCommentStore>(BugBashItemCommentStore);
    export var workItemFieldStore: WorkItemFieldStore = StoreFactory.getInstance<WorkItemFieldStore>(WorkItemFieldStore);
    export var workItemTypeStore: WorkItemTypeStore = StoreFactory.getInstance<WorkItemTypeStore>(WorkItemTypeStore);
    export var workItemTemplateStore: WorkItemTemplateStore = StoreFactory.getInstance<WorkItemTemplateStore>(WorkItemTemplateStore);
    export var workItemTemplateItemStore: WorkItemTemplateItemStore = StoreFactory.getInstance<WorkItemTemplateItemStore>(WorkItemTemplateItemStore);
}