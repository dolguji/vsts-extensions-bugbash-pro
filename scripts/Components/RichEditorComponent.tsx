import * as React from "react";

import { BaseComponent, IBaseComponentState } from "MB/Components/BaseComponent";
import { RichEditor, IRichEditorProps } from "MB/Components/RichEditor";
import { BaseStore } from "MB/Flux/Stores/BaseStore";
import { Loading } from "MB/Components/Loading";

import { StoresHub } from "../Stores/StoresHub";
import { SettingsActions } from "../Actions/SettingsActions";

import "../PasteImagePlugin";

export class RichEditorComponent extends BaseComponent<IRichEditorProps, IBaseComponentState> {
    protected getStores(): BaseStore<any, any, any>[] {
        return [StoresHub.bugBashSettingsStore];
    }

    protected initializeState() {
        this.state = {
            loading: true
        };
    }

    public componentDidMount() {
        super.componentDidMount();
        SettingsActions.initializeBugBashSettings();
    }

    protected getStoresState(): IBaseComponentState {                
        return {
            loading: StoresHub.bugBashSettingsStore.isLoading()
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