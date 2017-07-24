import "../../css/BugBashItemEditor.scss";

import * as React from "react";
import { autobind } from "OfficeFabric/Utilities";
import { Label } from "OfficeFabric/Label";
import { TextField } from "OfficeFabric/TextField";
import { Checkbox } from "OfficeFabric/Checkbox";
import { CommandBar } from "OfficeFabric/CommandBar";
import { IContextualMenuItem } from "OfficeFabric/components/ContextualMenu/ContextualMenu.Props";
import { ComboBox, IComboBoxOption } from "OfficeFabric/ComboBox";

import { MessagePanel, MessageType } from "VSTS_Extension/Components/Common/MessagePanel";
import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { Loading } from "VSTS_Extension/Components/Common/Loading";
import { InputError } from "VSTS_Extension/Components/Common/InputError";
import { BaseStore } from "VSTS_Extension/Flux/Stores/BaseStore";
import { IdentityView } from "VSTS_Extension/Components/WorkItemControls/IdentityView";

import { VersionControlChangeType, ItemContentType } from "TFS/VersionControl/Contracts";
import * as GitClient from "TFS/VersionControl/GitRestClient";
import Utils_Date = require("VSS/Utils/Date");
import Utils_Core = require("VSS/Utils/Core");
import { WebApiTeam } from "TFS/Core/Contracts";

import { RichEditorComponent } from "./RichEditorComponent";
import { buildGitPush, BugBashItemHelpers } from "../Helpers";
import { IBugBashItem, IBugBashItemComment } from "../Interfaces";
import { StoresHub } from "../Stores/StoresHub";
import { BugBashItemCommentActions } from "../Actions/BugBashItemCommentActions";

export interface IBugBashItemEditorProps extends IBaseComponentProps {
    bugBashItem: IBugBashItem;
    newComment?: string;
    onChange: (updatedBugBashItem: IBugBashItem, newComment?: string) => void;
    save: () => void;
}

export interface IBugBashItemEditorState extends IBaseComponentState {
    loadError?: string;
    error?: string;
    comments?: IBugBashItemComment[];
}

export class BugBashItemEditor extends BaseComponent<IBugBashItemEditorProps, IBugBashItemEditorState> {
    private _imagePastedHandler: (event, data) => void;

    constructor(props: IBugBashItemEditorProps, context?: any) {
        super(props, context);

        this._imagePastedHandler = Utils_Core.delegate(this, this._onImagePaste);
    }

    protected getStores(): BaseStore<any, any, any>[] {
        return [StoresHub.bugBashItemCommentStore, StoresHub.bugBashItemStore];
    }

    public componentDidMount(): void {
        super.componentDidMount();

        $(window).off("imagepasted", this._imagePastedHandler);
        $(window).on("imagepasted", this._imagePastedHandler);

        if (this.props.bugBashItem.id) {
             BugBashItemCommentActions.initializeComments(this.props.bugBashItem.id);
        } 
    }

    protected getStatesStore(): IBugBashItemEditorState {
        return {
            comments: StoresHub.bugBashItemCommentStore.getItem(this.props.bugBashItem.id)
        };
    }  

    protected initializeState() {
        this.state = {            
            comments: this.props.bugBashItem.id ? null : []            
        };
    }

    public componentWillReceiveProps(nextProps: Readonly<IBugBashItemEditorProps>): void {
        if (this.props.bugBashItem.id !== nextProps.bugBashItem.id) {
            if (nextProps.bugBashItem.id && StoresHub.bugBashItemCommentStore.isLoaded(nextProps.bugBashItem.id)) {
                this.updateState({
                    comments: StoresHub.bugBashItemCommentStore.getItem(nextProps.bugBashItem.id)
                });
            }
            else if (!nextProps.bugBashItem.id) {
                this.updateState({
                    comments: []
                });
            }
            else {
                this.updateState({
                    comments: null
                });
                BugBashItemCommentActions.initializeComments(nextProps.bugBashItem.id);
            }
        }
    }

    public componentWillUnmount() {
        $(window).off("imagepasted", this._imagePastedHandler);
    } 

    private _onChange(updatedBugBashItem: IBugBashItem, newComment?: string) {
        this.props.onChange(updatedBugBashItem, newComment);
    }

    @autobind
    private _onTitleChange(newValue: string) {
        let bugBashItem = {...this.props.bugBashItem};
        bugBashItem.title = newValue;
        this._onChange(bugBashItem);
    }

    @autobind
    private _onDescriptionChange(newValue: string) {
        let bugBashItem = {...this.props.bugBashItem};
        bugBashItem.description = newValue;
        this._onChange(bugBashItem);
    }

    @autobind
    private _onRejectionReasonChange(newValue: string) {
        let bugBashItem = {...this.props.bugBashItem};
        bugBashItem.rejectReason = newValue;
        this._onChange(bugBashItem);
    }

    @autobind
    private _onTeamChange(option: IComboBoxOption, _index: number, value: string) {
        let bugBashItem = {...this.props.bugBashItem};
        let newTeam: WebApiTeam = null;

        if (option) {
            const teamId = option.key as string;
            newTeam = StoresHub.teamStore.getItem(teamId);
        }
        
        bugBashItem.teamId = newTeam ? newTeam.id : value;
        this._onChange(bugBashItem);
    }

    @autobind
    private _onCommentChange(newValue: string) {
        let bugBashItem = {...this.props.bugBashItem};
        this._onChange(bugBashItem, newValue);
    }

    public render(): JSX.Element {
        const item = {...this.props.bugBashItem};

        if (this.state.loadError) {
            return <MessagePanel messageType={MessageType.Error} message={this.state.loadError} />;
        }
        else if (BugBashItemHelpers.isAccepted(item)) {
            return <MessagePanel 
                messageType={MessageType.Info} 
                message={"This item has been accepted. You can view or edit this item's work item from the \"Accepted items\" tab. Please refresh the list to clear this message."} />;
        }
        else {
            const allTeams = StoresHub.teamStore.getAll();
            const team = StoresHub.teamStore.getItem(item.teamId);

            return (
                <div className="item-editor" onKeyDown={this._onEditorKeyDown} tabIndex={0}>     
                    { this.state.error && <MessagePanel messageType={MessageType.Error} message={this.state.error} />}
                    
                    <TextField label="Title" 
                            value={item.title}
                            required={true} 
                            onGetErrorMessage={this._getTitleError}
                            onChanged={this._onTitleChange} />

                    <div className="item-team-container">
                        <Label required={true}>Team</Label>

                        <ComboBox                             
                            value={team ? team.name : item.teamId} 
                            allowFreeform={true}
                            autoComplete={"on"}
                            options={allTeams.map(t => {
                                return {
                                    key: t.id,
                                    text: t.name
                                };
                            })} 
                            onChanged={this._onTeamChange}/>
                        
                        { this._renderTeamError(item.teamId, team) }
                    </div>

                    { item.rejected && 
                        <TextField label="Reject reason" 
                                className="reject-reason-input"
                                value={item.rejectReason}
                                required={true} 
                                onGetErrorMessage={this._getRejectReasonError}
                                onChanged={this._onRejectionReasonChange} />
                    }

                    <div className="item-description-container">
                        <Label>Description</Label>
                        <RichEditorComponent 
                            containerId="description-editor" 
                            data={item.description} 
                            editorOptions={{
                                svgPath: `${VSS.getExtensionContext().baseUri}/css/libs/icons.svg`,
                                btns: [
                                    ['bold', 'italic'], 
                                    ['link'],
                                    ['insertImage'],
                                    'btnGrp-lists',
                                    ['removeformat'],
                                    ['fullscreen']
                                ]
                            }}
                            onChange={this._onDescriptionChange} />
                    </div>

                    <div className="item-comments-editor-container">
                        <Label>Discussion</Label>
                        <RichEditorComponent 
                            containerId="comment-editor" 
                            data={this.props.newComment || ""} 
                            editorOptions={{
                                svgPath: `${VSS.getExtensionContext().baseUri}/css/libs/icons.svg`,
                                btns: []
                            }}
                            onChange={this._onCommentChange} />
                    </div>

                    <div className="item-comments-container">
                        {this._renderComments()}
                    </div>
                </div>
            );
        }        
    }

    private _renderComments(): React.ReactNode {
        if (this.props.bugBashItem.id) {
            let comments = StoresHub.bugBashItemCommentStore.getItem(this.props.bugBashItem.id);
            if (!comments) {
                return <Loading />;
            }
            else {
                comments = comments.slice();
                comments = comments.sort((c1: IBugBashItemComment, c2: IBugBashItemComment) => {
                    return -1 * Utils_Date.defaultComparer(c1.createdDate, c2.createdDate);
                });
                
                return comments.map((comment: IBugBashItemComment, index: number) => {
                    return (
                        <div className="item-comment" key={`${index}`}>
                            <div className="created-by">
                                <IdentityView identityDistinctName={comment.createdBy} />
                                <div className="created-date">{Utils_Date.friendly(comment.createdDate)}</div>
                            </div>
                            <div className="message" dangerouslySetInnerHTML={{ __html: comment.content }} />                        
                        </div>
                    );
                });
            }
        }
        else {
            return null;
        }
    }

    @autobind
    private _onEditorKeyDown(e: React.KeyboardEvent<any>) {
        if (e.ctrlKey && e.keyCode === 83) {
            e.preventDefault();
            this.props.save();
        }
    }    

    private _renderTeamError(teamId: string, team: WebApiTeam): JSX.Element {
        if (teamId == null || teamId.trim() === "") {
            return <InputError error={`Team is required.`} />;
        }
        else if (team == null) {
            return <InputError error={`${teamId} team does not exist.`} />;
        }
    }

    @autobind
    private _getTitleError(value: string): string | IPromise<string> {
        if (value == null || value.trim() === "") {
            return "Title is required";
        }
        if (value.length > 256) {
            return `The length of the title should less than 256 characters, actual is ${value.length}.`;
        }
        return "";
    }

    @autobind
    private _getRejectReasonError(value: string): string | IPromise<string> {
        if (value == null || value.trim().length == 0) {
            return "Reject reason can not be empty";
        }

        if (value.length > 128) {
            return `The length of the reject reason should less than 128 characters, actual is ${value.length}.`;
        }
        return "";
    }

    private async _onImagePaste(_event, args) {
        const data = args.data;
        const callback = args.callback;

        const settings = StoresHub.settingsStore.getAll();
        if (settings && settings.gitMediaRepo) {
            const dataStartIndex = data.indexOf(",") + 1;
            const metaPart = data.substring(5, dataStartIndex - 1);
            const dataPart = data.substring(dataStartIndex);

            const extension = metaPart.split(";")[0].split("/").pop();
            const fileName = `pastedImage_${Date.now().toString()}.${extension}`;
            const gitPath = `BugBash_${StoresHub.bugBashStore.getItem(this.props.bugBashItem.bugBashId).title.replace(" ", "_")}/pastedImages/${fileName}`;
            const projectId = VSS.getWebContext().project.id;

            try {
                const gitClient = GitClient.getClient();
                const gitItem = await gitClient.getItem(settings.gitMediaRepo, "/", projectId);
                const pushModel = buildGitPush(gitPath, gitItem.commitId, VersionControlChangeType.Add, dataPart, ItemContentType.Base64Encoded);
                await gitClient.createPush(pushModel, settings.gitMediaRepo, projectId);

                const imageUrl = `${VSS.getWebContext().collection.uri}/${VSS.getWebContext().project.id}/_api/_versioncontrol/itemContent?repositoryId=${settings.gitMediaRepo}&path=${gitPath}&version=GBmaster&contentOnly=true`;
                callback(imageUrl);
            }
            catch (e) {
                callback(null);
            }
        }
        else {
            callback(null);
        }
    }   
}