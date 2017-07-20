import { StoresHub } from "../Stores/StoresHub";
import { WorkItemActionsHub } from "./ActionsHub";
import { WorkItemStore } from "../Stores/WorkItemStore";

import { JsonPatchDocument, JsonPatchOperation, Operation } from "VSS/WebApi/Contracts";
import * as WitClient from "TFS/WorkItemTracking/RestClient";
import { WorkItem, WorkItemErrorPolicy } from "TFS/WorkItemTracking/Contracts";

export module WorkItemActions {
    export async function initializeWorkItems(ids: number[], fields?: string[]) {
        if (!ids || ids.length === 0) {
            WorkItemActionsHub.AddOrUpdateWorkItems.invoke(null);
        }
        else if (!StoresHub.workItemStore.isLoading()) {
            const idsToFetch: number[] = [];
            for (const id of ids) {
                if (!StoresHub.workItemStore.isLoaded(id)) {
                    idsToFetch.push(id);
                }
            }

            if (idsToFetch.length == 0) {
                WorkItemActionsHub.AddOrUpdateWorkItems.invoke(null);
            }
            
            StoresHub.workItemStore.setLoading(true);

            try {
                const workItems = await WitClient.getClient().getWorkItems(idsToFetch, fields, null, null, WorkItemErrorPolicy.Omit);
                WorkItemActionsHub.AddOrUpdateWorkItems.invoke(workItems.filter(w => w != null));
                StoresHub.workItemStore.setLoading(false);
            }
            catch (e) {
                StoresHub.workItemStore.setLoading(false);
                throw e.message;
            }  
        }   
    }

    export async function refreshWorkItems(ids: number[], fields?: string[]) {
        if (!ids || ids.length === 0) {
            WorkItemActionsHub.AddOrUpdateWorkItems.invoke(null);
        }
        else if (!StoresHub.workItemStore.isLoading()) {
            StoresHub.workItemStore.setLoading(true);

            try {
                const workItems = await WitClient.getClient().getWorkItems(ids, fields, null, null, WorkItemErrorPolicy.Omit);
                WorkItemActionsHub.AddOrUpdateWorkItems.invoke(workItems.filter(w => w != null));
                StoresHub.workItemStore.setLoading(false);
            }
            catch (e) {
                StoresHub.workItemStore.setLoading(false);
                throw e.message;
            }  
        }   
    }

    export async function initializeWorkItem(id: number, fields?: string[]) {
        if (!StoresHub.workItemStore.isLoaded(id)) {
            WorkItemActionsHub.AddOrUpdateWorkItems.invoke(null);
        }
        else if (!StoresHub.workItemStore.isLoading()) {
            StoresHub.workItemStore.setLoading(true);

            try {
                const workItem = await WitClient.getClient().getWorkItem(id, fields);
                WorkItemActionsHub.AddOrUpdateWorkItems.invoke([workItem]);
                StoresHub.workItemStore.setLoading(false);
            }
            catch (e) {
                StoresHub.workItemStore.setLoading(false);
                throw e.message;
            }  
        }   
    }

    export async function refreshWorkItem(id: number, fields?: string[]) {        
        if (!StoresHub.workItemStore.isLoading()) {
            StoresHub.workItemStore.setLoading(true);

            try {
                const workItem = await WitClient.getClient().getWorkItem(id, fields);
                WorkItemActionsHub.AddOrUpdateWorkItems.invoke([workItem]);
                StoresHub.workItemStore.setLoading(false);
            }
            catch (e) {
                StoresHub.workItemStore.setLoading(false);
                throw e.message;
            }  
        }
    }

    export async function createWorkItem(workItemType: string, fieldValues: IDictionaryStringTo<string>): Promise<WorkItem> {
        if (!StoresHub.workItemStore.isLoading()) {
            StoresHub.workItemStore.setLoading(true);
            
            let patchDocument: JsonPatchDocument & JsonPatchOperation[] = [];
            for (let fieldRefName in fieldValues) {
                patchDocument.push({
                    op: Operation.Add,
                    path: `/fields/${fieldRefName}`,
                    value: fieldValues[fieldRefName]
                } as JsonPatchOperation);
            }

            try {
                const workItem = await WitClient.getClient().createWorkItem(patchDocument, VSS.getWebContext().project.id, workItemType);
                WorkItemActionsHub.AddOrUpdateWorkItems.invoke([workItem]);
                StoresHub.workItemStore.setLoading(false);
                return workItem;
            }
            catch (e) {
                StoresHub.workItemStore.setLoading(false);
                throw e.message;
            }
        }
    }

    export async function updateWorkItem(workItemId: number, fieldValues: IDictionaryStringTo<string>): Promise<WorkItem> {
        if (!StoresHub.workItemStore.isLoading()) {
            StoresHub.workItemStore.setLoading(true);
            
            let patchDocument: JsonPatchDocument & JsonPatchOperation[] = [];
            for (let fieldRefName in fieldValues) {
                patchDocument.push({
                    op: Operation.Add,
                    path: `/fields/${fieldRefName}`,
                    value: fieldValues[fieldRefName]
                } as JsonPatchOperation);
            }

            try {
                const workItem = await WitClient.getClient().updateWorkItem(patchDocument, workItemId);
                WorkItemActionsHub.AddOrUpdateWorkItems.invoke([workItem]);
                StoresHub.workItemStore.setLoading(false);
                return workItem;
            }
            catch (e) {
                StoresHub.workItemStore.setLoading(false);
                throw e.message;
            }
        }
    }
}