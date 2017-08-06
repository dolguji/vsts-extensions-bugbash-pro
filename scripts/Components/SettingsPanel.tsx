import "../../css/SettingsPanel.scss";

import * as React from "react";
import { GitRepository } from "TFS/VersionControl/Contracts";
import Utils_String = require("VSS/Utils/String");

import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { BaseStore } from "VSTS_Extension/Flux/Stores/BaseStore";
import { GitRepoActions } from "VSTS_Extension/Flux/Actions/GitRepoActions";
import { Loading } from "VSTS_Extension/Components/Common/Loading";
import { InfoLabel } from "VSTS_Extension/Components/Common/InfoLabel";

import { autobind } from "OfficeFabric/Utilities";
import { Label } from "OfficeFabric/Label";
import { Dropdown, IDropdownOption, IDropdownProps } from "OfficeFabric/Dropdown";
import { PrimaryButton } from "OfficeFabric/Button";

import { StoresHub } from "../Stores/StoresHub";
import { UserSettings, BugBashSettings } from "../Interfaces";
import { SettingsActions } from "../Actions/SettingsActions";

interface ISettingsPanelState extends IBaseComponentState {
    origBugBashSettings?: BugBashSettings;
    newBugBashSettings?: BugBashSettings;
    origUserSettings?: UserSettings;
    newUserSettings?: UserSettings;
    gitRepos?: GitRepository[];
}

export class SettingsPanel extends BaseComponent<IBaseComponentProps, ISettingsPanelState> {
    protected getStores(): BaseStore<any, any, any>[] {
        return [StoresHub.bugBashSettingsStore, StoresHub.gitRepoStore];
    }

    protected initializeState() {
        this.state = {
            loading: true,
            gitRepos: []
        };
    }

    public componentDidMount() {
        super.componentDidMount();
        SettingsActions.initializeBugBashSettings();
        GitRepoActions.initializeGitRepos();
    }

    protected getStoresState(): ISettingsPanelState {
        return {
            newBugBashSettings: {...StoresHub.bugBashSettingsStore.getAll()},
            origBugBashSettings: {...StoresHub.bugBashSettingsStore.getAll()},
            gitRepos: StoresHub.gitRepoStore.getAll(),
            loading: StoresHub.bugBashSettingsStore.isLoading() || StoresHub.gitRepoStore.isLoading()
        };
    }

    public render(): JSX.Element {
        if (this.state.loading) {
            return <Loading />;
        }
        else {
            let gitDropdownOptions = this.state.gitRepos.map((repo: GitRepository, index: number) => {
                return {
                    key: repo.id,
                    index: index,
                    text: repo.name,
                    selected: Utils_String.equals(this.state.newBugBashSettings.gitMediaRepo, repo.id, true)
                };
            });

            return <div className="settings-panel">                
                <div className="settings-controls-container">
                    <Label className="settings-label">Project Settings</Label>
                    <div className="settings-control">
                        <InfoLabel label="Media Git Repo" info="Select a git repo to store media and attachments" />
                        <Dropdown                 
                            className="git-repo-dropdown"
                            onRenderList={this._onRenderCallout} 
                            options={gitDropdownOptions} 
                            onChanged={(option: IDropdownOption) => {
                                let newSettings = {...this.state.newBugBashSettings};
                                newSettings.gitMediaRepo = option.key as string;
                                this.updateState({newBugBashSettings: newSettings});
                            }} /> 
                    </div>
                    <PrimaryButton className="save-button" disabled={!this._isSettingsDirty()} onClick={this._onSaveClick}>
                        Save
                    </PrimaryButton>
                </div>

                <div className="settings-controls-container">
                    <Label className="settings-label">User Settings</Label>
                    <div className="settings-control">
                        <InfoLabel label="Media Git Repo" info="Select a git repo to store media and attachments" />
                        <Dropdown                 
                            className="git-repo-dropdown"
                            onRenderList={this._onRenderCallout} 
                            options={gitDropdownOptions} 
                            onChanged={(option: IDropdownOption) => {
                                let newSettings = {...this.state.newBugBashSettings};
                                newSettings.gitMediaRepo = option.key as string;
                                this.updateState({newBugBashSettings: newSettings});
                            }} /> 
                    </div>
                    <PrimaryButton className="save-button" disabled={!this._isSettingsDirty()} onClick={this._onSaveClick}>
                        Save
                    </PrimaryButton>
                </div>               
            </div>;
        }        
    }

    private _isSettingsDirty(): boolean {
        return this.state.newBugBashSettings.gitMediaRepo !== this.state.origBugBashSettings.gitMediaRepo;
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
            SettingsActions.updateBugBashSettings(this.state.newBugBashSettings);
        }        
    }
}