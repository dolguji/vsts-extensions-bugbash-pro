import * as React from "react";

import { MessagePanel, MessageType } from "VSTS_Extension/Components/Common/MessagePanel";
import { Loading } from "VSTS_Extension/Components/Common/Loading";
import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { BaseStore } from "VSTS_Extension/Stores/BaseStore";

import Utils_String = require("VSS/Utils/String");

import { IBugBash } from "../Interfaces";
import { BugBashEditor } from "./BugBashEditor";
import { BugBashStore } from "../Stores/BugBashStore";
import { StoresHub } from "../Stores/StoresHub";

interface IEditBugBashViewState extends IBaseComponentState {
    item: IBugBash;
    loading: boolean;
}

interface IEditBugBashViewProps extends IBaseComponentState {
    id: string;
}

export class EditBugBashView extends BaseComponent<IEditBugBashViewProps, IEditBugBashViewState> {
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
                return <MessagePanel messageType={MessageType.Error} message="This instance of bug bash doesn't exist." />;
            }
            else if(!Utils_String.equals(VSS.getWebContext().project.id, this.state.item.projectId, true)) {
                return <MessagePanel messageType={MessageType.Error} message="This instance of bug bash is out of scope of current project." />;
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