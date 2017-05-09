import * as React from "react";

import { MessageBar, MessageBarType } from "OfficeFabric/MessageBar";

import { Loading } from "VSTS_Extension/Components/Common/Loading";
import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { BaseStore } from "VSTS_Extension/Stores/BaseStore";

import Utils_String = require("VSS/Utils/String");

import { IBugBash } from "../Models";
import { BugBashEditor } from "./BugBashEditor";
import { BugBashStore } from "../Stores/BugBashStore";
import { StoresHub } from "../Stores/StoresHub";

interface IEditHubViewState extends IBaseComponentState {
    item: IBugBash;
    loading: boolean;
}

interface IEditHubViewProps extends IBaseComponentState {
    id: string;
}

export class EditBugBashView extends BaseComponent<IEditHubViewProps, IEditHubViewState> {
    protected initializeState() {
        this.state = {
            item: null,
            loading: true
        };
    }

    protected getStoresToLoad(): {new (): BaseStore<any, any, any>}[] {
        return [BugBashStore];
    }

    protected async initialize() {
        let found = await StoresHub.bugBashStore.ensureItem(this.props.id);
        if (!found) {
            this.updateState({
                item: null,
                loading: false
            });
        }
    }

    public render(): JSX.Element {
        if (this.state.loading) {
            return <Loading />;
        }
        else {
            if (!this.state.item) {
                return <MessageBar messageBarType={MessageBarType.error}>This instance of bug bash doesn't exist.</MessageBar>;
            }
            else if(!Utils_String.equals(VSS.getWebContext().project.id, this.state.item.projectId, true)) {
                return <MessageBar messageBarType={MessageBarType.error}>This instance of bug bash is out of scope of current project.</MessageBar>;
            }
            else {
                return <BugBashEditor id={this.state.item.id} />;
            }            
        }
    }    

    protected onStoreChanged() {
        this.updateState({
            item: StoresHub.bugBashStore.getItem(this.props.id),
            loading: StoresHub.bugBashStore.isLoaded() ? false : true
        });
    }
}