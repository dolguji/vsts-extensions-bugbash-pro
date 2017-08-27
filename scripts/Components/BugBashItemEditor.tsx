import "../../css/BugBashItemEditor.scss";

import * as React from "react";
import { autobind } from "OfficeFabric/Utilities";
import { Label } from "OfficeFabric/Label";
import { TextField } from "OfficeFabric/TextField";
import { Overlay } from "OfficeFabric/Overlay";
import { MessageBar, MessageBarType } from "OfficeFabric/MessageBar";

import { ComboBox } from "VSTS_Extension/Components/Common/Combo/Combobox";
import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { Loading } from "VSTS_Extension/Components/Common/Loading";
import { InputError } from "VSTS_Extension/Components/Common/InputError";
import { BaseStore } from "VSTS_Extension/Flux/Stores/BaseStore";
import { IdentityView } from "VSTS_Extension/Components/WorkItemControls/IdentityView";

import { delegate, delay, DelayedFunction } from "VSS/Utils/Core";
import { VersionControlChangeType, ItemContentType } from "TFS/VersionControl/Contracts";
import * as GitClient from "TFS/VersionControl/GitRestClient";
import Utils_Date = require("VSS/Utils/Date");
import { WebApiTeam } from "TFS/Core/Contracts";

import { RichEditorComponent } from "./RichEditorComponent";
import { buildGitPush } from "../Helpers";
import { IBugBashItem, IBugBashItemComment } from "../Interfaces";
import { StoresHub } from "../Stores/StoresHub";
import { BugBashItemCommentActions } from "../Actions/BugBashItemCommentActions";
import { BugBashFieldNames } from "../Constants";
import { BugBashItem } from "../ViewModels/BugBashItem";

export interface IBugBashItemEditorProps extends IBaseComponentProps {
    bugBashItem: BugBashItem;    
}

export interface IBugBashItemEditorState extends IBaseComponentState {    
    comments?: IBugBashItemComment[];
    commentsLoading?: boolean;
}

export class BugBashItemEditor extends BaseComponent<IBugBashItemEditorProps, IBugBashItemEditorState> {
    private _imagePastedHandler: (event, data) => void;
    private _updateBugBashDelayedFunction: DelayedFunction;

    constructor(props: IBugBashItemEditorProps, context?: any) {
        super(props, context);

        this._imagePastedHandler = delegate(this, this._onImagePaste);
    }

    protected getStores(): BaseStore<any, any, any>[] {
        return [StoresHub.bugBashItemCommentStore, StoresHub.bugBashItemStore];
    }

    public componentDidMount() {
        super.componentDidMount();

        $(window).off("imagepasted", this._imagePastedHandler);
        $(window).on("imagepasted", this._imagePastedHandler);

        if (this.props.bugBashItem.id) {
             BugBashItemCommentActions.initializeComments(this.props.bugBashItem.id);
        } 
    }

    protected getStoresState(): IBugBashItemEditorState {
        return {
            comments: StoresHub.bugBashItemCommentStore.getItem(this.props.bugBashItem.id),
            commentsLoading: StoresHub.bugBashItemCommentStore.isLoading(this.props.bugBashItem.id),
            loading: StoresHub.bugBashItemStore.isLoading(this.props.bugBashItem.id)
        };
    }  

    protected initializeState() {
        this.state = {            
            comments: this.props.bugBashItem.id ? null : [],
            commentsLoading: this.props.bugBashItem.id ? true : false
        };
    }

    public componentWillReceiveProps(nextProps: Readonly<IBugBashItemEditorProps>) {
        if (this.props.bugBashItem.id !== nextProps.bugBashItem.id) {
            if (nextProps.bugBashItem.id && StoresHub.bugBashItemCommentStore.isLoaded(nextProps.bugBashItem.id)) {
                this.updateState({
                    comments: StoresHub.bugBashItemCommentStore.getItem(nextProps.bugBashItem.id),
                    commentsLoading: false
                });
            }
            else if (!nextProps.bugBashItem.id) {
                this.updateState({
                    comments: [],
                    commentsLoading: false
                });
            }
            else {
                this.updateState({
                    comments: null,
                    commentsLoading: true
                });
                BugBashItemCommentActions.initializeComments(nextProps.bugBashItem.id);
            }
        }
    }

    public componentWillUnmount() {
        $(window).off("imagepasted", this._imagePastedHandler);
    } 

    private _onChange(updatedBugBashItem: IBugBashItem, newComment?: string) {
        if (this._updateBugBashDelayedFunction) {
            this._updateBugBashDelayedFunction.cancel();
        }

        this._updateBugBashDelayedFunction = delay(this, 200, this.props.onChange, [updatedBugBashItem, newComment]);
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
    private _onTeamChange(newTeamName: string) {
        let bugBashItem = {...this.props.bugBashItem};
        const newTeam = StoresHub.teamStore.getItem(newTeamName);
        bugBashItem.teamId = newTeam ? newTeam.id : newTeamName;
        this._onChange(bugBashItem);
    }

    @autobind
    private _onCommentChange(newValue: string) {
        let bugBashItem = {...this.props.bugBashItem};
        this._onChange(bugBashItem, newValue);
    }

    public render(): JSX.Element {
        const item = {...this.props.bugBashItem};
        
        if (BugBashItemHelpers.isAccepted(item)) {
            const webContext = VSS.getWebContext();
            const witUrl = `${webContext.collection.uri}/${webContext.project.name}/_workitems/edit/${item.workItemId}`;

            return <div className="item-editor accepted-item">
                <MessageBar messageBarType={MessageBarType.success} className="message-panel">
                    Accepted: <a href={witUrl} target="_blank">Work item: {item.workItemId}</a>
                </MessageBar>
            </div>;
        }

        const allTeams = StoresHub.teamStore.getAll();
        const team = StoresHub.teamStore.getItem(item.teamId);

        return (
            <div className="item-editor" onKeyDown={this._onEditorKeyDown} tabIndex={0}>                    
                { this.state.loading && <Overlay className="loading-overlay"><Loading /></Overlay>}
                { this.props.error && <MessageBar messageBarType={MessageBarType.error} className="message-panel">{this.props.error}</MessageBar> }

                <TextField label="Title" 
                    value={item.title}
                    required={true} 
                    onChanged={this._onTitleChange} />

                { this._renderTitleError(item.title) }

                <div className="item-team-container">
                    <Label required={true}>Team</Label>

                    <ComboBox                             
                        value={team ? team.name : item.teamId} 
                        options={{
                            type: "list",
                            mode: "drop",
                            allowEdit: true,
                            source: allTeams.map(t => t.name)
                        }} 
                        error={this._getTeamError(item.teamId, team)}
                        onChange={this._onTeamChange}/>                    
                </div>

                { item.rejected && 
                    <TextField label="Reject reason" 
                            className="reject-reason-input"
                            value={item.rejectReason}
                            required={true} 
                            onChanged={this._onRejectionReasonChange} />                        
                }

                { item.rejected && this._renderRejectReasonError(item.rejectReason) }

                <div className="item-description-container">
                    <Label>Description</Label>
                    <RichEditorComponent 
                        containerId="description-editor" 
                        data={item.description} 
                        editorOptions={{
                            svgPath: `${VSS.getExtensionContext().baseUri}/css/libs/icons.png`,
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
                            svgPath: `${VSS.getExtensionContext().baseUri}/css/libs/icons.png`,
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

    private _renderComments(): React.ReactNode {
        if (this.props.bugBashItem.id) {            
            if (this.state.commentsLoading || !this.state.comments) {
                return <Loading />;
            }
            else {
                let comments = this.state.comments.slice();
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

    private _getTeamError(teamId: string, team: WebApiTeam): string {
        if (teamId == null || teamId.trim() === "") {
            return "Team is required.";
        }
        else if (team == null) {
            return `${teamId} team does not exist.`
        }

        return null;
    }

    private _renderTitleError(title: string): JSX.Element {
        if (title == null || title.trim() === "") {
            return <InputError error="Title is required." />;
        }
        else if (title.length > 256) {
            return <InputError error={`The length of the title should be less than 257 characters, actual is ${title.length}.`} />;
        }

        return null;
    }

    @autobind
    private _renderRejectReasonError(value: string): JSX.Element {
        if (value == null || value.trim().length == 0) {
            return <InputError error="Reject reason can not be empty." />;
        }
        else if (value.length > 128) {
            return <InputError error={`The length of the reject reason should less than 129 characters, actual is ${value.length}.`} />;
        }

        return null;
    }

    private async _onImagePaste(_event, args) {
        const data = args.data;
        const callback = args.callback;

        const settings = StoresHub.bugBashSettingsStore.getAll();
        if (settings && settings.gitMediaRepo) {
            const dataStartIndex = data.indexOf(",") + 1;
            const metaPart = data.substring(5, dataStartIndex - 1);
            const dataPart = data.substring(dataStartIndex);

            const extension = metaPart.split(";")[0].split("/").pop();
            const fileName = `pastedImage_${Date.now().toString()}.${extension}`;
            const gitPath = `BugBash_${StoresHub.bugBashStore.getItem(this.props.bugBashItem.bugBashId).getFieldValue<string>(BugBashFieldNames.Title, true).replace(" ", "_")}/pastedImages/${fileName}`;
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