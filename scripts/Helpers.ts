import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");
import {JsonPatchDocument, JsonPatchOperation, Operation} from "VSS/WebApi/Contracts";
import * as WitClient from "TFS/WorkItemTracking/RestClient";
import { WorkItem } from "TFS/WorkItemTracking/Contracts";
import * as WitBatchClient from "TFS/WorkItemTracking/BatchRestClient";

export async function saveWorkItems(fieldValuesMap: IDictionaryNumberTo<IDictionaryStringTo<string>>): Promise<WorkItem[]> {
    let batchDocument: [number, JsonPatchDocument][] = [];

    for (const id in fieldValuesMap) {
        let patchDocument: JsonPatchDocument & JsonPatchOperation[] = [];
        for (let fieldRefName in fieldValuesMap[id]) {
            patchDocument.push({
                op: Operation.Add,
                path: `/fields/${fieldRefName}`,
                value: fieldValuesMap[id][fieldRefName]
            } as JsonPatchOperation);
        }

        batchDocument.push([parseInt(id), patchDocument]);
    }

    let response = await WitBatchClient.getClient().updateWorkItemsBatch(batchDocument);
    return response.value.map((v: WitBatchClient.JsonHttpResponse) => JSON.parse(v.body) as WorkItem);
}

export async function saveWorkItem(id: number, fieldValues: IDictionaryStringTo<string>): Promise<WorkItem> {
    let patchDocument: JsonPatchDocument & JsonPatchOperation[] = [];
    for (let fieldRefName in fieldValues) {
        patchDocument.push({
            op: Operation.Add,
            path: `/fields/${fieldRefName}`,
            value: fieldValues[fieldRefName]
        } as JsonPatchOperation);
    }

    return await WitClient.getClient().updateWorkItem(patchDocument, id);
}

export async function createWorkItem(workItemType: string, fieldValues: IDictionaryStringTo<string>): Promise<WorkItem> {
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

export function isInteger(value: string): boolean {
    return /^\d+$/.test(value);
}

export function parseTags(tags: string): string[] {
    if (tags && tags.trim()) {
        let tagsArr = (tags || "").split(";");
        return tagsArr.map((t: string) => t.trim());
    }
    return [];
}
