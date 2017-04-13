import * as React from "react";

import { IBugBash, LoadingState } from "../Models";
import { HubView, IHubViewState } from "./HubView";
import { Loading } from "./Loading";
import { MessagePanel, MessageType } from "./MessagePanel";
import { BugBashEditor } from "./BugBashEditor";

interface IEditHubViewState extends IHubViewState {
    item: IBugBash;
}

export class EditBugBashView extends HubView<IEditHubViewState> {    
    public render(): JSX.Element {
        if (this.state.loadingState === LoadingState.Loading) {
            return <Loading />;
        }
        else {
            if (!this.state.item) {
                return <MessagePanel message="This instance of bug bash doesnt exist in the context of current team." messageType={MessageType.Error} />
            }
            else {
                return <BugBashEditor id={this.state.item.id} context={this.props.context} />;
            }            
        }
    }

    protected async initialize() {
        let found = await this.props.context.actionsCreator.ensureBugBash(this.props.id);
        if (!found) {
            this.setState({
                item: null,
                loadingState: LoadingState.Loaded
            });
        }
    }

    protected getStateFromStore(): IEditHubViewState {
        const item = this.props.context.stores.bugBashItemStore.getItem(this.props.id);
        return {
            item: item,
            loadingState: this.props.context.stores.bugBashItemStore.isLoaded() ? LoadingState.Loaded : LoadingState.Loading
        };
    }
}