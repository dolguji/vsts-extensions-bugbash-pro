import * as React from "react";
import { autobind } from "OfficeFabric/Utilities";
import { Label } from "OfficeFabric/Label";
import { TextField } from "OfficeFabric/TextField";
import { Button, ButtonType } from "OfficeFabric/Button";

import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";

import { WorkItemTemplate, WorkItem, FieldType } from "TFS/WorkItemTracking/Contracts";
import { WorkItemFormNavigationService } from "TFS/WorkItemTracking/Services";

import { IBugBashItemDocument } from "../Models";
import Helpers = require("../Helpers");

export interface IBugBashItemProps extends IBaseComponentProps {
    item: IBugBashItemDocument;
}

export interface IBugBashItemState extends IBaseComponentState {
    
}

export class BugBashItem extends BaseComponent<IBugBashItemProps, IBugBashItemState> {
      
}