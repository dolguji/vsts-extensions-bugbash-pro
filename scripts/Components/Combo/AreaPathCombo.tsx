private _getTreeNodes(node: WorkItemClassificationNode, uiNode: TreeNode, level: number): TreeNode {
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
                this._getTreeNodes(node, uiNode, level + 1);
            }
        }
        return uiNode;
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