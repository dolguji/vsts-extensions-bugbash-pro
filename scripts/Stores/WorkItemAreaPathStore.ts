import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");
import { WorkItemClassificationNode, TreeStructureGroup } from "TFS/WorkItemTracking/Contracts";
import * as WitClient from "TFS/WorkItemTracking/RestClient";

import { BaseStore } from "VSTS_Extension/Stores/BaseStore";

export class WorkItemAreaPathStore extends BaseStore<IDictionaryStringTo<WorkItemClassificationNode>, WorkItemClassificationNode, string> {
    constructor() {
        super();
        this.items = {};    
    }

    protected getItemByKey(projectId: string): WorkItemClassificationNode {  
        return this.items[projectId.toLowerCase()];
    }

    protected async initializeItems(): Promise<void> {
        
    }

    public async ensureAreaPathNode(projectId?: string): Promise<boolean> {
        if (!this.itemExists(projectId)) {
            try {
                let node = await WitClient.getClient().getClassificationNode(VSS.getWebContext().project.id || projectId, TreeStructureGroup.Areas, null, 5);
                if (node) {
                    this._onAdd(node, VSS.getWebContext().project.id || projectId);
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

    public getKey(): string {
        return "WorkItemAreaPathStore";
    }

    private _onAdd(item: WorkItemClassificationNode, projectId: string): void {
        if (!item || !projectId) {
            return;
        }

        if (!this.items) {
            this.items = {};
        }

        this.items[projectId] = item;

        this.emitChanged();
    }
}