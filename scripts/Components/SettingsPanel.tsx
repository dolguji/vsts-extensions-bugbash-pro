import "../../css/SettingsPanel.scss";

import * as React from "react";
import { GitRepository } from "TFS/VersionControl/Contracts";
import { WebApiTeam } from "TFS/Core/Contracts";

import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "MB/Components/BaseComponent";
import { BaseStore } from "MB/Flux/Stores/BaseStore";
import { GitRepoActions } from "MB/Flux/Actions/GitRepoActions";
import { TeamActions } from "MB/Flux/Actions/TeamActions";
import { Loading } from "MB/Components/Loading";
import { InfoLabel } from "MB/Components/InfoLabel";

import { autobind } from "OfficeFabric/Utilities";
import { Dropdown, IDropdownOption, IDropdownProps } from "OfficeFabric/Dropdown";
import { Label } from "OfficeFabric/Label";
import { PrimaryButton } from "OfficeFabric/Button";
import { MessageBar, MessageBarType } from "OfficeFabric/MessageBar";

import { StoresHub } from "../Stores/StoresHub";
import { IUserSettings, IBugBashSettings } from "../Interfaces";
import { SettingsActions } from "../Actions/SettingsActions";
import { ErrorKeys } from "../Constants";
import { BugBashErrorMessageActions } from "../Actions/BugBashErrorMessageActions";

interface ISettingsPanelState extends IBaseComponentState {
    origBugBashSettings: IBugBashSettings;
    newBugBashSettings: IBugBashSettings;
    origUserSettings: IUserSettings;
    newUserSettings: IUserSettings;
    gitRepos: IDropdownOption[];
    teams: IDropdownOption[];
    error?: string;
}

export class SettingsPanel extends BaseComponent<IBaseComponentProps, ISettingsPanelState> {
    protected getStores(): BaseStore<any, any, any>[] {
        return [StoresHub.bugBashSettingsStore, StoresHub.gitRepoStore, StoresHub.teamStore, StoresHub.userSettingsStore, StoresHub.bugBashErrorMessageStore];
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
        const isLoading = StoresHub.bugBashSettingsStore.isLoading() || StoresHub.gitRepoStore.isLoading() || StoresHub.teamStore.isLoading() || StoresHub.userSettingsStore.isLoading();
        const bugBashSettings = StoresHub.bugBashSettingsStore.getAll();
        let userSetting = StoresHub.userSettingsStore.getItem(VSS.getWebContext().user.email);

        if (isLoading) {
            userSetting = null;
        }
        else {
            if (userSetting == null) {
                userSetting = {
                    id: VSS.getWebContext().user.email,
                    __etag: 0,
                    associatedTeam: null
                }
            }
        }

        let state = {
            newBugBashSettings: bugBashSettings ? {...bugBashSettings} : null,
            origBugBashSettings: bugBashSettings ? {...bugBashSettings} : null,
            newUserSettings: userSetting ? {...userSetting} : null,
            origUserSettings: userSetting ? {...userSetting} : null,
            loading: isLoading,
            error: StoresHub.bugBashErrorMessageStore.getItem(ErrorKeys.BugBashSettingsError)
        } as ISettingsPanelState;

        if (StoresHub.teamStore.isLoaded() && !this.state.teams) {
            const emptyItem = [
                {   
                    key: "", text: "<No team>"
                }
            ];

            const teams: IDropdownOption[] = emptyItem.concat(StoresHub.teamStore.getAll().map((t: WebApiTeam) => {
                return {
                    key: t.id,
                    text: t.name
                }
            }));

            state = {...state, teams: teams} as ISettingsPanelState;
        }

        if (StoresHub.gitRepoStore.isLoaded() && !this.state.gitRepos) {
            const emptyItem = [
                {   
                    key: "", text: "<No repo>"
                }
            ];

            const repos: IDropdownOption[] = emptyItem.concat(StoresHub.gitRepoStore.getAll().map((r: GitRepository) => {
                return {
                    key: r.id,
                    text: r.name
                }
            }));

            state = {...state, gitRepos: repos} as ISettingsPanelState;
        }

        return state;
    }

    public render(): JSX.Element {
        if (this.state.loading) {
            return <Loading />;
        }
        else {
            return <div className="settings-panel">  
                { this.state.error && 
                    <MessageBar 
                        className="message-panel"
                        messageBarType={MessageBarType.error} 
                        onDismiss={this._dismissErrorMessage}>
                        {"Settings could not be saved due to an unknown error. Please refresh the page and try again."}
                    </MessageBar>
                }              
                <div className="settings-controls-container">
                    <Label className="settings-label">Project Settings</Label>
                    <div className="settings-control">
                        <InfoLabel label="Media Git Repo" info="Select a git repo to store media and attachments" />
                        <Dropdown
                            options={this.state.gitRepos}
                            onRenderList={this._onRenderCallout}
                            selectedKey={this.state.newBugBashSettings.gitMediaRepo}
                            onChanged={(option?: IDropdownOption) => {
                                let newSettings = {...this.state.newBugBashSettings};
                                newSettings.gitMediaRepo = option.key as string;
                                this.updateState({newBugBashSettings: newSettings} as ISettingsPanelState);
                            }}
                        />
                    </div>
                    <PrimaryButton className="save-button" disabled={!this._isSettingsDirty()} onClick={this._onSaveClick}>
                        Save
                    </PrimaryButton>
                </div>

                <div className="settings-controls-container">
                    <Label className="settings-label">User Settings</Label>
                    <div className="settings-control">
                        <InfoLabel label="Associated team" info="Select a team associated with you." />
                        <Dropdown
                            options={this.state.teams}
                            onRenderList={this._onRenderCallout}
                            selectedKey={this.state.newUserSettings.associatedTeam}
                            onChanged={(option?: IDropdownOption) => {
                                let newSettings = {...this.state.newUserSettings};
                                newSettings.associatedTeam = option.key as string;
                                this.updateState({newUserSettings: newSettings} as ISettingsPanelState);
                            }}
                        />
                    </div>
                    <PrimaryButton className="save-button" disabled={!this._isUserSettingsDirty()} onClick={this._onSaveUserSettingClick}>
                        Save
                    </PrimaryButton>
                </div>               
            </div>;
        }        
    }

    @autobind
    private _dismissErrorMessage() {
        BugBashErrorMessageActions.dismissErrorMessage(ErrorKeys.BugBashSettingsError);
    };

    private _isSettingsDirty(): boolean {
        return this.state.newBugBashSettings.gitMediaRepo !== this.state.origBugBashSettings.gitMediaRepo;
    }

    private _isUserSettingsDirty(): boolean {
        return this.state.newUserSettings.associatedTeam !== this.state.origUserSettings.associatedTeam;
    }

    @autobind
    private async _onSaveClick(): Promise<void> {
        if (this._isSettingsDirty()) {
            SettingsActions.updateBugBashSettings(this.state.newBugBashSettings);
        }        
    }

    @autobind
    private async _onSaveUserSettingClick(): Promise<void> {
        if (this._isUserSettingsDirty()) {
            SettingsActions.updateUserSettings(this.state.newUserSettings);
        }        
    }

    @autobind
    private _onRenderCallout(props?: IDropdownProps, defaultRender?: (props?: IDropdownProps) => JSX.Element): JSX.Element {
        return <div className="callout-container">
                {defaultRender(props)}
            </div>;        
    }
}