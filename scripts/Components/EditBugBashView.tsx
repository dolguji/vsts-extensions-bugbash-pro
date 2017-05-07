import * as React from "react";

import { MessageBar, MessageBarType } from 'OfficeFabric/MessageBar';

import { StoreFactory } from "VSTS_Extension/Stores/BaseStore";
import { Loading } from "VSTS_Extension/Components/Common/Loading";

import { IBugBash } from "../Models";
import { HubView, IHubViewState } from "./HubView";
import { BugBashEditor } from "./BugBashEditor";

interface IEditHubViewState extends IHubViewState {
    item: IBugBash;
}

export class EditBugBashView extends HubView<IEditHubViewState> {    
    public render(): JSX.Element {
        if (this.state.loading) {
            return <Loading />;
        }
        else {
            if (!this.state.item) {
                return <MessageBar messageBarType={MessageBarType.error}>This instance of bug bash doesnt exist in the context of current project.</MessageBar>
            }
            else {
                return <BugBashEditor id={this.state.item.id} />;
            }            
        }
    }

    protected async initialize() {
        let found = await this.bugBashItemStore.ensureBugBash(this.props.id);
        if (!found) {
            this.setState({
                item: null,
                loading: false
            });
        }
    }

    protected getStateFromStore(): IEditHubViewState {
        const item = this.bugBashItemStore.getItem(this.props.id);
        return {
            item: item,
            loading: this.bugBashItemStore.isLoaded() ? false : true
        };
    }
}