import * as React from "react";

import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { BugBashEditor } from "./BugBashEditor";

export class NewBugBashView extends BaseComponent<IBaseComponentProps, IBaseComponentState> {    
    public render(): JSX.Element {
        return <BugBashEditor />;
    }
}