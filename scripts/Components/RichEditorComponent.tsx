import * as React from "react";

import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { RichEditor, IRichEditorProps } from "VSTS_Extension/Components/Common/RichEditor/RichEditor";
import { BaseStore } from "VSTS_Extension/Stores/BaseStore";
import { Loading } from "VSTS_Extension/Components/Common/Loading";

import { StoresHub } from "../Stores/StoresHub";
import { SettingsStore } from "../Stores/SettingsStore";
import { Settings } from "../Interfaces";

import "../PasteImagePlugin";

interface IRichComponentState extends IBaseComponentState {
    loading?: boolean;
}

export class RichEditorComponent extends BaseComponent<IRichEditorProps, IRichComponentState> {
    protected getStoresToLoad(): {new (): BaseStore<any, any, any>}[] {
        return [SettingsStore];
    }

    protected initializeState() {
        this.state = {
            loading: true
        };
    }

    protected initialize(): void {
        StoresHub.settingsStore.initialize();
    }

    protected onStoreChanged() {                
        this.updateState({
            loading: !StoresHub.settingsStore.isLoaded()
        });
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