import * as React from "react";

import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { RichEditor, IRichEditorProps } from "VSTS_Extension/Components/Common/RichEditor/RichEditor";
import { BaseStore } from "VSTS_Extension/Flux/Stores/BaseStore";
import { Loading } from "VSTS_Extension/Components/Common/Loading";

import { StoresHub } from "../Stores/StoresHub";
import { SettingsActions } from "../Actions/SettingsActions";

import "../PasteImagePlugin";

interface IRichComponentState extends IBaseComponentState {
    loading?: boolean;
}

export class RichEditorComponent extends BaseComponent<IRichEditorProps, IRichComponentState> {
    protected getStores(): BaseStore<any, any, any>[] {
        return [StoresHub.settingsStore];
    }

    protected initializeState() {
        this.state = {
            loading: true
        };
    }

    public componentDidMount(): void {
        super.componentDidMount();
        SettingsActions.initializeBugBashSettings();
    }

    protected getStoresState(): IRichComponentState {                
        return {
            loading: StoresHub.settingsStore.isLoading()
        };
    }

    public render(): JSX.Element {
        if (this.state.loading) {
            return <Loading />;
        }
        else {
            return <RichEditor {...this.props} />;
        }  
    }
}