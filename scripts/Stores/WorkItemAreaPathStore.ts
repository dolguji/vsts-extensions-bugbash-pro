import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");
import { WorkItemClassificationNode, TreeStructureGroup } from "TFS/WorkItemTracking/Contracts";
import * as WitClient from "TFS/WorkItemTracking/RestClient";

import { BaseStore } from "VSTS_Extension/Stores/BaseStore";

export class WorkItemAreaPathStore extends BaseStore<IDictionaryStringTo<WorkItemClassificationNode>, WorkItemClassificationNode, string> {
    private _allowedAreaPaths: IDictionaryStringTo<string[]>;

    constructor() {
        super();
        this.items = {};
        this._allowedAreaPaths = {};
    }

    protected getItemByKey(projectId: string): WorkItemClassificationNode {  
        return this.items[(projectId || VSS.getWebContext().project.id).toLowerCase()];
    }

    protected async initializeItems(): Promise<void> {
        
    }

    public getAreaPaths(projectId?: string): string[] {
        return this._allowedAreaPaths[(projectId || VSS.getWebContext().project.id).toLowerCase()];
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
        if (!this._allowedAreaPaths) {
            this._allowedAreaPaths = {};
        }

        this.items[projectId.toLowerCase()] = item;
        this._allowedAreaPaths[projectId.toLowerCase()] = this._getNodePaths(item);

        this.emitChanged();
    }

    private _getNodePaths(node: WorkItemClassificationNode, parentNodeName?: string): string[] {
        let nodeName = parentNodeName ? `${parentNodeName}\\${node.name}`: node.name;
        let returnData: string[] = [nodeName];
        if (node.children) {
            for (let child of node.children) {
                returnData = returnData.concat(this._getNodePaths(child, nodeName));
            }
        }

        return returnData;        
    }
}