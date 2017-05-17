/// <reference types="react" />

import * as React from "react";

import { ComboBox } from "./ComboBox";
import { StoresHub } from "../../Stores/StoresHub";
import { WorkItemAreaPathStore } from "../../Stores/WorkItemAreaPathStore";

import {TreeNode} from "VSS/Controls/TreeView";
import { WorkItemClassificationNode } from "TFS/WorkItemTracking/Contracts";
import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { BaseStore } from "VSTS_Extension/Stores/BaseStore";

export interface IAreaPathComboProps extends IBaseComponentProps {
    value?: string;
    onChange: (newValue: string) => void;
}

export interface IAreaPathComboState extends IBaseComponentState {
    treeNodes?: TreeNode;
}

export class AreaPathCombo extends BaseComponent<IAreaPathComboProps, IAreaPathComboState> {
    protected getDefaultClassName(): string {
        return "areapathcombo";
    }

    protected getStoresToLoad(): {new (): BaseStore<any, any, any>}[] {
        return [WorkItemAreaPathStore];
    }

    protected initialize(): void {
        StoresHub.workItemAreaPathStore.ensureAreaPathNode();
    }

    protected onStoreChanged() {
        const rootNode = StoresHub.workItemAreaPathStore.getItem(VSS.getWebContext().project.id);

        if (rootNode) {
            this.updateState({
                treeNodes: this._getTreeNode(rootNode, null, 1)
            });
        }        
    }

    protected initializeState() {
        this.state = {};
    }

    public render(): JSX.Element {
        if (this.state.treeNodes) {
            return <ComboBox 
                value={this.props.value} 
                onChange={this.props.onChange}
                options={{
                    type: "treeSearch",
                    mode: "drop",
                    initialLevel: 2,
                    sepChar: "\\",
                    source: [this.state.treeNodes],
                    enabled: true,
                    allowEdit: true
                }} />;
        }

        return null;
    }

    private _getTreeNode(node: WorkItemClassificationNode, uiNode: TreeNode, level: number): TreeNode {
        let nodes = node.children;
        let newUINode: TreeNode;
        let nodeName = node.name;

        level = level || 1;
        if (uiNode) {
            newUINode = TreeNode.create(nodeName);
            uiNode.add(newUINode);
            uiNode = newUINode;
        }
        else {
            uiNode = TreeNode.create(nodeName);
        }
        uiNode.expanded = level < 2;
        if (nodes) {
            for (let node of nodes) {
                this._getTreeNode(node, uiNode, level + 1);
            }
        }
        return uiNode;
    }
}
