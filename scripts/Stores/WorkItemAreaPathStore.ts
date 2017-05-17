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
        return this.items[(projectId || VSS.getWebContext().project.id).toLowerCase()];
    }

    protected async initializeItems(): Promise<void> {
        
    }

    public async ensureAreaPathNode(projectId?: string): Promise<boolean> {
        if (!this.itemExists(projectId)) {
            try {
                let node = await WitClient.getClient().getClassificationNode(projectId || VSS.getWebContext().project.id, TreeStructureGroup.Areas, null, 5);
                if (node) {
                    this._onAdd(node, projectId || VSS.getWebContext().project.id);
                    return true;
                }
                else {
                    return false;
                }
            }
            catch (e) {
                return false;
            }
        }
        else {
            this.emitChanged();
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

        this.items[projectId.toLowerCase()] = item;

        this.emitChanged();
    }
}