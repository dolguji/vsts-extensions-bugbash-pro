import "../../css/SettingsPanel.scss";

import * as React from "react";
import { GitRepository } from "TFS/VersionControl/Contracts";
import Utils_String = require("VSS/Utils/String");
import { WebApiTeam } from "TFS/Core/Contracts";

import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { BaseStore } from "VSTS_Extension/Flux/Stores/BaseStore";
import { GitRepoActions } from "VSTS_Extension/Flux/Actions/GitRepoActions";
import { TeamActions } from "VSTS_Extension/Flux/Actions/TeamActions";
import { Loading } from "VSTS_Extension/Components/Common/Loading";
import { InfoLabel } from "VSTS_Extension/Components/Common/InfoLabel";

import { autobind } from "OfficeFabric/Utilities";
import { ComboBox, IComboBoxOption } from "OfficeFabric/Combobox";
import { Label } from "OfficeFabric/Label";
import { Dropdown, IDropdownOption, IDropdownProps } from "OfficeFabric/Dropdown";
import { PrimaryButton } from "OfficeFabric/Button";

import { StoresHub } from "../Stores/StoresHub";
import { IUserSettings, IBugBashSettings } from "../Interfaces";
import { SettingsActions } from "../Actions/SettingsActions";

interface ISettingsPanelState extends IBaseComponentState {
    origBugBashSettings: IBugBashSettings;
    newBugBashSettings: IBugBashSettings;
    origUserSettings: IUserSettings;
    newUserSettings: IUserSettings;
    gitRepos: GitRepository[];
    teams: WebApiTeam[];
}

export class SettingsPanel extends BaseComponent<IBaseComponentProps, ISettingsPanelState> {
    protected getStores(): BaseStore<any, any, any>[] {
        return [StoresHub.bugBashSettingsStore, StoresHub.gitRepoStore, StoresHub.teamStore, StoresHub.userSettingsStore];
    }

    protected initializeState() {
        this.state = {
            loading: true,
        } as ISettingsPanelState;
    }

    public componentDidMount() {
        super.componentDidMount();
        SettingsActions.initializeBugBashSettings();
        SettingsActions.initializeUserSettings();
        TeamActions.initializeTeams();
        GitRepoActions.initializeGitRepos();
    }

    protected getStoresState(): ISettingsPanelState {
        return {
            newBugBashSettings: {...StoresHub.bugBashSettingsStore.getAll()},
            origBugBashSettings: {...StoresHub.bugBashSettingsStore.getAll()},
            newUserSettings: {...StoresHub.userSettingsStore.getItem(VSS.getWebContext().user.email)},
            origUserSettings: {...StoresHub.userSettingsStore.getItem(VSS.getWebContext().user.email)},
            gitRepos: StoresHub.gitRepoStore.getAll(),
            teams: StoresHub.teamStore.getAll(),
            loading: StoresHub.bugBashSettingsStore.isLoading() || StoresHub.gitRepoStore.isLoading() || StoresHub.teamStore.isLoading() || StoresHub.userSettingsStore.isLoading()
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
                                this.updateState({newBugBashSettings: newSettings} as ISettingsPanelState);
                            }} /> 
                    </div>
                    <PrimaryButton className="save-button" disabled={!this._isSettingsDirty()} onClick={this._onSaveClick}>
                        Save
                    </PrimaryButton>
                </div>

                <div className="settings-controls-container">
                    <Label className="settings-label">User Settings</Label>
                    <div className="settings-control">
                        <InfoLabel label="Associated team" info="Select a team associated with you." />
                        <ComboBox
                            options={this.state.teams.map(t => {
                                return {
                                    key: t.id,
                                    text: t.name
                                }
                            })}
                            allowFreeform={true}
                            autoComplete={"on"}
                            selectedKey={this.state.newUserSettings.associatedTeam}
                            onChanged={(option?: IComboBoxOption) => {
                                let newSettings = {...this.state.newUserSettings};
                                newSettings.associatedTeam = option.key as string;
                                this.updateState({newUserSettings: newSettings} as ISettingsPanelState);
                            }}
                        />
                    </div>
                    <PrimaryButton className="save-button" disabled={!this._isUserSettingsDirty()} onClick={this._onSaveClick}>
                        Save
                    </PrimaryButton>
                </div>               
            </div>;
        }        
    }

    private _isSettingsDirty(): boolean {
        return this.state.newBugBashSettings.gitMediaRepo !== this.state.origBugBashSettings.gitMediaRepo;
    }

    private _isUserSettingsDirty(): boolean {
        return this.state.newUserSettings.associatedTeam !== this.state.origUserSettings.associatedTeam;
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