import { StoresHub } from "../Stores/StoresHub";
import { WorkItemActionsHub } from "./ActionsHub";
import { WorkItemStore } from "../Stores/WorkItemStore";

import { JsonPatchDocument, JsonPatchOperation, Operation } from "VSS/WebApi/Contracts";
import * as WitClient from "TFS/WorkItemTracking/RestClient";
import { WorkItem } from "TFS/WorkItemTracking/Contracts";

export module WorkItemActions {
    export async function initializeWorkItems(ids: number[], fields?: string[]) {
        const idsToFetch: number[] = [];
        for (const id of ids) {
            if (!StoresHub.workItemStore.itemExists(id)) {
                idsToFetch.push(id);
            }
        }

        const workItems = await WitClient.getClient().getWorkItems(idsToFetch, fields);
        StoresHub.workItemStore.
    }

    async function createWorkItem(workItemType: string, fieldValues: IDictionaryStringTo<string>): Promise<WorkItem> {
        let patchDocument: JsonPatchDocument & JsonPatchOperation[] = [];
        for (let fieldRefName in fieldValues) {
            patchDocument.push({
                op: Operation.Add,
                path: `/fields/${fieldRefName}`,
                value: fieldValues[fieldRefName]
            } as JsonPatchOperation);
        }

        return await WitClient.getClient().createWorkItem(patchDocument, VSS.getWebContext().project.id, workItemType);
    }

    async function updateWorkItem(workItemId: number, fieldValues: IDictionaryStringTo<string>): Promise<WorkItem> {
        let patchDocument: JsonPatchDocument & JsonPatchOperation[] = [];
        for (let fieldRefName in fieldValues) {
            patchDocument.push({
                op: Operation.Add,
                path: `/fields/${fieldRefName}`,
                value: fieldValues[fieldRefName]
            } as JsonPatchOperation);
        }

        return await WitClient.getClient().updateWorkItem(patchDocument, workItemId);
    }
}