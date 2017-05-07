import * as React from "react";

import { HubView, IHubViewState } from "./HubView";
import { BugBashEditor } from "./BugBashEditor";

export class NewBugBashView extends HubView<IHubViewState> {    
    public render(): JSX.Element {
        return <BugBashEditor />;
    }

    protected getStateFromStore(): IHubViewState {
        return {
            loading: true
        };
    }
}