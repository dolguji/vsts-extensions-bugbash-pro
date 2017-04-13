import * as React from "react";

import { LoadingState } from "../Models";
import { HubView, IHubViewState } from "./HubView";
import { Loading } from "./Loading";
import { BugBashEditor } from "./BugBashEditor";

interface INewHubViewState extends IHubViewState {

}

export class NewBugBashView extends HubView<INewHubViewState> {    
    public render(): JSX.Element {
        return <BugBashEditor context={this.props.context} />;
    }

    protected initialize() {
        
    }

    protected getStateFromStore(): INewHubViewState {
        return {
            loadingState: LoadingState.Loaded
        };
    }
}