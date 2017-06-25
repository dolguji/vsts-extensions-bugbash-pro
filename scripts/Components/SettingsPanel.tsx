import "../../css/SettingsPanel.scss";

import * as React from "react";
import { GitRepository } from "TFS/VersionControl/Contracts";
import Utils_String = require("VSS/Utils/String");

import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { BaseStore } from "VSTS_Extension/Stores/BaseStore";
import { Loading } from "VSTS_Extension/Components/Common/Loading";
import { InfoLabel } from "VSTS_Extension/Components/Common/InfoLabel";

import { autobind } from "OfficeFabric/Utilities";
import { Label } from "OfficeFabric/Label";
import { Dropdown, IDropdownOption, IDropdownProps } from "OfficeFabric/Dropdown";
import { PrimaryButton } from "OfficeFabric/Button";

import { StoresHub } from "../Stores/StoresHub";
import { SettingsStore } from "../Stores/SettingsStore";
import { GitRepoStore } from "../Stores/GitRepoStore";
import { Settings } from "../Interfaces";

interface ISettingsPanelState extends IBaseComponentState {
    loading?: boolean;
    origSettings?: Settings;
    newSettings?: Settings;
}

export class SettingsPanel extends BaseComponent<IBaseComponentProps, ISettingsPanelState> {
    protected getStoresToLoad(): {new (): BaseStore<any, any, any>}[] {
        return [SettingsStore, GitRepoStore];
    }

    protected initializeState() {
        this.state = {
            loading: true
        };
    }

    protected initialize(): void {
        StoresHub.settingsStore.initialize();
        StoresHub.gitRepoStore.initialize();
    }

    protected onStoreChanged() {                
        this.updateState({
            newSettings: {...StoresHub.settingsStore.getAll()},
            origSettings: {...StoresHub.settingsStore.getAll()},
            loading: !StoresHub.settingsStore.isLoaded() || !StoresHub.gitRepoStore.isLoaded()
        });
    }

    public render(): JSX.Element {
        if (this.state.loading) {
            return <Loading />;
        }
        else {
            let gitDropdownOptions = StoresHub.gitRepoStore.getAll().map((repo: GitRepository, index: number) => {
                return {
                    key: repo.id,
                    index: index,
                    text: repo.name,
                    selected: Utils_String.equals(this.state.newSettings.gitMediaRepo, repo.id, true)
                };
            });

            return <div className="settings-panel">
                <Label className="settings-label">Settings</Label>
                <div className="settings-controls">
                    <div className="settings-control-container">
                        <InfoLabel label="Media Git Repo" info="Select a git repo to store media and attachments" />
                        <Dropdown                 
                            className="git-repo-dropdown"
                            onRenderList={this._onRenderCallout} 
                            options={gitDropdownOptions} 
                            onChanged={(option: IDropdownOption) => {
                                let newSettings = {...this.state.newSettings};
                                newSettings.gitMediaRepo = option.key as string;
                                this.updateState({newSettings: newSettings});
                            }} /> 
                    </div>
                </div>

                <PrimaryButton className="save-button" disabled={!this._isSettingsDirty()} onClick={this._onSaveClick}>
                    Save
                </PrimaryButton>
            </div>;
        }        
    }

    private _isSettingsDirty(): boolean {
        return this.state.newSettings.gitMediaRepo !== this.state.origSettings.gitMediaRepo;
    }

    @autobind
    private _onRenderCallout(props?: IDropdownProps, defaultRender?: (props?: IDropdownProps) => JSX.Element): JSX.Element {
        return (
            <div className="callout-container">
                {defaultRender(props)}
            </div>
        );
    }

    @autobind
    private async _onSaveClick(): Promise<void> {
        if (this._isSettingsDirty()) {
            StoresHub.settingsStore.updateSettings(this.state.newSettings);
        }        
    }
}