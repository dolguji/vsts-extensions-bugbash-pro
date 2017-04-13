import { Action } from "VSS/Flux/Action";
import { WorkItemField, WorkItemTemplateReference, WorkItemType, WorkItemTemplate } from "TFS/WorkItemTracking/Contracts";

import { IBugBash } from "../Models";
import { IBugBashItemStore } from "../Stores/BugbashItemStore";
import { IWorkItemFieldStore } from "../Stores/WorkItemFieldStore";
import { IWorkItemTypeStore } from "../Stores/WorkItemTypeStore";
import { IWorkItemTemplateStore } from "../Stores/WorkItemTemplateStore";
import { IWorkItemTemplateItemStore } from "../Stores/WorkItemTemplateItemStore";
import { BugBashManager } from "../BugbashManager";

import * as WitClient from "TFS/WorkItemTracking/RestClient";

export class ActionsHub {
    public InitializeBugBashItems = new Action<IBugBash[]>();
    public ClearBugBashItems = new Action<void>();
    public BugBashItemDeleted = new Action<IBugBash>();
    public BugBashItemAdded = new Action<IBugBash>();
    public BugBashItemUpdated = new Action<IBugBash>();
    
    public InitializeWorkItemFields = new Action<WorkItemField[]>();
    public InitializeWorkItemTemplates = new Action<WorkItemTemplateReference[]>();
    public InitializeWorkItemTypes = new Action<WorkItemType[]>();

    public WorkItemTemplateItemAdded = new Action<WorkItemTemplate | WorkItemTemplate[]>();
}

export class ActionsCreator {
    constructor(
        private _actionsHub: ActionsHub, 
        private _bugBashItemDataProvider: IBugBashItemStore, 
        private _workItemFieldDataProvider: IWorkItemFieldStore,
        private _workItemTemplateDataProvider: IWorkItemTemplateStore,
        private _workItemTypeDataProvider: IWorkItemTypeStore,
        private _workItemTemplateItemDataProvider: IWorkItemTemplateItemStore) {
    }  

    public async initializeAllBugBashes() {
        if (this._bugBashItemDataProvider.isLoaded()) {
            // Do nothing if query hierarchy data is already loaded
            this._actionsHub.InitializeBugBashItems.invoke(null);
        }
        else {            
            let bugbashes = await BugBashManager.readBugBashes();
            this._actionsHub.InitializeBugBashItems.invoke(bugbashes);
        }
    }

    public async refreshAllBugBashes() {
        if (this._bugBashItemDataProvider.isLoaded()) {
            // refresh only if the store has already been laoded
            this._actionsHub.ClearBugBashItems.invoke(null);

            let bugbashes = await BugBashManager.readBugBashes();
            this._actionsHub.InitializeBugBashItems.invoke(bugbashes);
        }
    }

    public async initializeWorkItemFields() {
        if (this._workItemFieldDataProvider.isLoaded()) {
            // Do nothing if query hierarchy data is already loaded
            this._actionsHub.InitializeWorkItemFields.invoke(null);
        }
        else {            
            let fields = await WitClient.getClient().getFields(VSS.getWebContext().project.id);
            this._actionsHub.InitializeWorkItemFields.invoke(fields);
        }
    }

    public async initializeWorkItemTemplates() {
        if (this._workItemTemplateDataProvider.isLoaded()) {
            // Do nothing if query hierarchy data is already loaded
            this._actionsHub.InitializeWorkItemTemplates.invoke(null);
        }
        else {            
            let templates = await WitClient.getClient().getTemplates(VSS.getWebContext().project.id, VSS.getWebContext().team.id);
            this._actionsHub.InitializeWorkItemTemplates.invoke(templates);
        }
    }

    public async initializeWorkItemTypes() {
        if (this._workItemTypeDataProvider.isLoaded()) {
            // Do nothing if query hierarchy data is already loaded
            this._actionsHub.InitializeWorkItemTypes.invoke(null);
        }
        else {            
            let workItemTypes = await WitClient.getClient().getWorkItemTypes(VSS.getWebContext().project.id);
            this._actionsHub.InitializeWorkItemTypes.invoke(workItemTypes);
        }
    }
    
    public async ensureBugBash(id: string): Promise<boolean> {
        if (!this._bugBashItemDataProvider.itemExists(id)) {
            try {
                let bugbash = await BugBashManager.readBugBash(id);
                if (bugbash) {
                    this._actionsHub.BugBashItemAdded.invoke(bugbash);
                    return true;
                }
            }
            catch (e) {
                return false;
            }

            return false;
        }
        else {
            return true;
        }
    }  

    public async ensureTemplateItem(id: string): Promise<boolean> {
        if (!this._workItemTemplateItemDataProvider.itemExists(id)) {
            try {
                let template = await WitClient.getClient().getTemplate(VSS.getWebContext().project.id, VSS.getWebContext().team.id, id)
                if (template) {
                    this._actionsHub.WorkItemTemplateItemAdded.invoke(template);
                    return true;
                }
            }
            catch (e) {
                return false;
            }

            return false;
        }
        else {
            return true;
        }
    }

    public async deleteBugBash(bugBash: IBugBash): Promise<boolean> {         
        let deleted = await BugBashManager.deleteBugBash(bugBash);

        if (deleted) {
            this._actionsHub.BugBashItemDeleted.invoke(bugBash);
        }

        return deleted;
    }

    public async createBugBash(bugBash: IBugBash): Promise<IBugBash> {     
        bugBash.id = bugBash.id || Date.now().toString();    
        let savedBugBash = await BugBashManager.writeBugBash(bugBash);

        if (savedBugBash) {
            this._actionsHub.BugBashItemAdded.invoke(savedBugBash);
        }

        return savedBugBash;
    }

    public async updateBugBash(bugBash: IBugBash): Promise<IBugBash> {         
        let savedBugBash = await BugBashManager.writeBugBash(bugBash);

        if (savedBugBash) {
            this._actionsHub.BugBashItemUpdated.invoke(savedBugBash);
        }

        return savedBugBash;
    }
}
