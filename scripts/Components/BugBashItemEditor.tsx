import "../../css/BugBashItemEditor.scss";

import * as React from "react";
import { autobind } from "OfficeFabric/Utilities";
import { Label } from "OfficeFabric/Label";
import { TextField } from "OfficeFabric/TextField";
import { Overlay } from "OfficeFabric/Overlay";
import { MessageBar, MessageBarType } from "OfficeFabric/MessageBar";

import { ComboBox } from "MB/Components/Combobox";
import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "MB/Components/BaseComponent";
import { Loading } from "MB/Components/Loading";
import { InputError } from "MB/Components/InputError";
import { BaseStore } from "MB/Flux/Stores/BaseStore";
import { IdentityView } from "MB/Components/IdentityView";
import { DateUtils } from "MB/Utils/Date";
import { CoreUtils } from "MB/Utils/Core";

import { WebApiTeam } from "TFS/Core/Contracts";

import { RichEditorComponent } from "./RichEditorComponent";
import { copyImageToGitRepo } from "../Helpers";
import { IBugBashItemComment } from "../Interfaces";
import { StoresHub } from "../Stores/StoresHub";
import { BugBashItemCommentActions } from "../Actions/BugBashItemCommentActions";
import { BugBashItemFieldNames, ErrorKeys, BugBashFieldNames, SizeLimits } from "../Constants";
import { BugBashItem } from "../ViewModels/BugBashItem";
import { BugBashErrorMessageActions } from "../Actions/BugBashErrorMessageActions";

export interface IBugBashItemEditorProps extends IBaseComponentProps {
    bugBashItem: BugBashItem;
    bugBashId: string;
}

export interface IBugBashItemEditorState extends IBaseComponentState {    
    comments?: IBugBashItemComment[];
    commentsLoading?: boolean;
    error?: string;
}

export class BugBashItemEditor extends BaseComponent<IBugBashItemEditorProps, IBugBashItemEditorState> {
    private _teamNames: string[];
    private _imagePastedHandler: (event, data) => void;
    private _updateBugBashItemDelayedFunction: CoreUtils.DelayedFunction;

    constructor(props: IBugBashItemEditorProps, context?: any) {
        super(props, context);

        this._imagePastedHandler = CoreUtils.delegate(this, this._onImagePaste);
        this._teamNames = StoresHub.teamStore.getAll().map(t => t.name);        
    }

    protected getStores(): BaseStore<any, any, any>[] {
        return [StoresHub.bugBashItemCommentStore, StoresHub.bugBashItemStore, StoresHub.bugBashErrorMessageStore];
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
            loading: StoresHub.bugBashItemStore.isLoading(this.props.bugBashItem.id),
            error: StoresHub.bugBashErrorMessageStore.getItem(ErrorKeys.BugBashItemError)
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
            CoreUtils.delay(this, 0, () => {
                this._dismissErrorMessage();
            });

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
        super.componentWillUnmount();
        $(window).off("imagepasted", this._imagePastedHandler);
        this._dismissErrorMessage();
    }

    public render(): JSX.Element {
        const item = this.props.bugBashItem;
        
        if (item.isAccepted) {
            const webContext = VSS.getWebContext();

            const witUrl = `${webContext.collection.uri}${webContext.project.name}/_workitems/edit/${item.workItemId}`;

            return <div className="item-editor accepted-item">
                <MessageBar messageBarType={MessageBarType.success} className="message-panel">
                    Accepted: <a href={witUrl} target="_blank">Work item: {item.workItemId}</a>
                </MessageBar>
            </div>;
        }
        
        const teamId = item.getFieldValue<string>(BugBashItemFieldNames.TeamId);
        const title = item.getFieldValue<string>(BugBashItemFieldNames.Title);
        const description = item.getFieldValue<string>(BugBashItemFieldNames.Description);
        const rejected = item.getFieldValue<boolean>(BugBashItemFieldNames.Rejected);
        const rejectReason = item.getFieldValue<string>(BugBashItemFieldNames.RejectReason);

        const team = StoresHub.teamStore.getItem(teamId);

        return (
            <div className="item-editor" onKeyDown={this._onEditorKeyDown} tabIndex={0}>                    
                { this.state.loading && <Overlay className="loading-overlay"><Loading /></Overlay>}
                { this.state.error && <MessageBar messageBarType={MessageBarType.error} onDismiss={this._dismissErrorMessage} className="message-panel">{this.state.error}</MessageBar> }

                <TextField label="Title" 
                    maxLength={SizeLimits.TitleFieldMaxLength}
                    value={title}
                    required={true} 
                    onChanged={(newValue: string) => this._onChange(BugBashItemFieldNames.Title, newValue)} />

                { this._renderTitleError(title) }

                <div className="item-team-container">
                    <Label required={true}>Team</Label>

                    <ComboBox                             
                        value={team ? team.name : teamId} 
                        options={{
                            type: "list",
                            mode: "drop",
                            allowEdit: true,
                            source: this._teamNames
                        }} 
                        error={this._getTeamError(teamId, team)}
                        onChange={this._onTeamChange}/>
                </div>

                { rejected && 
                    <TextField label="Reject reason" 
                        maxLength={SizeLimits.RejectFieldMaxLength}
                        className="reject-reason-input"
                        value={rejectReason}
                        required={true} 
                        onChanged={(newValue: string) => this._onChange(BugBashItemFieldNames.RejectReason, newValue)} />                        
                }

                { rejected && this._renderRejectReasonError(rejectReason) }

                <div className="item-description-container">
                    <Label>Description</Label>
                    <RichEditorComponent 
                        containerId="description-editor" 
                        data={description} 
                        editorOptions={{
                            svgPath: `${VSS.getExtensionContext().baseUri}/css/libs/icons.png`,
                            btns: [
                                ['bold', 'italic'], 
                                ['link'],
                                ['insertImage'],
                                ['upload'],
                                'btnGrp-lists',
                                ['removeformat'],
                                ['fullscreen']
                            ]
                        }}
                        onChange={(newValue: string) => this._onChange(BugBashItemFieldNames.Description, newValue)} />
                </div>

                <div className="item-comments-editor-container">
                    <Label>Discussion</Label>
                    <RichEditorComponent 
                        containerId="comment-editor" 
                        data={item.newComment || ""} 
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
                    return -1 * DateUtils.defaultComparer(c1.createdDate, c2.createdDate);
                });
                
                return comments.map((comment: IBugBashItemComment, index: number) => {
                    return (
                        <div className="item-comment" key={`${index}`}>
                            <div className="created-by">
                                <IdentityView identityDistinctName={comment.createdBy} />
                                <div className="created-date">{DateUtils.friendly(comment.createdDate)}</div>
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

    private _getTeamError(teamId: string, team: WebApiTeam): string {
        if (teamId == null || teamId.trim() === "") {
            return "Team is required.";
        }
        else if (team == null) {
            return `"${teamId}" team does not exist.`
        }

        return null;
    }

    private _renderTitleError(title: string): JSX.Element {
        if (title == null || title.trim() === "") {
            return <InputError error="Title is required." />;
        }
        else if (title.length > SizeLimits.TitleFieldMaxLength) {
            return <InputError error={`The length of the title should be less than 257 characters, actual is ${title.length}.`} />;
        }

        return null;
    }

    @autobind
    private _renderRejectReasonError(value: string): JSX.Element {
        if (value == null || value.trim().length == 0) {
            return <InputError error="Reject reason can not be empty." />;
        }
        else if (value.length > SizeLimits.RejectFieldMaxLength) {
            return <InputError error={`The length of the reject reason should less than 129 characters, actual is ${value.length}.`} />;
        }

        return null;
    }

    private async _onImagePaste(_event, args) {
        const gitPath = StoresHub.bugBashStore.getItem(this.props.bugBashId).getFieldValue<string>(BugBashFieldNames.Title, true).replace(" ", "_");

        try {
            const imageUrl = await copyImageToGitRepo(args.data, gitPath);
            args.callback(imageUrl);
        }
        catch (e) {
            BugBashErrorMessageActions.showErrorMessage(e, ErrorKeys.BugBashItemError);
        }
    }   

    @autobind
    private _onEditorKeyDown(e: React.KeyboardEvent<any>) {
        if (e.ctrlKey && e.keyCode === 83) {
            e.preventDefault();
            this.props.bugBashItem.save(this.props.bugBashId);
        }
    }

    private _onChange<T>(fieldName: BugBashItemFieldNames, fieldValue: T, immediate?: boolean) {
        if (this._updateBugBashItemDelayedFunction) {
            this._updateBugBashItemDelayedFunction.cancel();
        }
        let bugBashItem = this.props.bugBashItem;

        if (immediate) {
            bugBashItem.setFieldValue<T>(fieldName, fieldValue);
        }
        else {
            this._updateBugBashItemDelayedFunction = CoreUtils.delay(this, 200, () => {
                bugBashItem.setFieldValue<T>(fieldName, fieldValue);
            });
        }
    } 

    @autobind
    private _onCommentChange(newComment: string) {
        if (this._updateBugBashItemDelayedFunction) {
            this._updateBugBashItemDelayedFunction.cancel();
        }
        let bugBashItem = this.props.bugBashItem;

        this._updateBugBashItemDelayedFunction = CoreUtils.delay(this, 200, () => {
            bugBashItem.setComment(newComment);
        });        
    }

    @autobind
    private _onTeamChange(newTeamName: string) {
        const newTeam = StoresHub.teamStore.getItem(newTeamName);
        const teamId = newTeam ? newTeam.id : newTeamName;
        this._onChange(BugBashItemFieldNames.TeamId, teamId);
    }

    @autobind
    private _dismissErrorMessage() {
        BugBashErrorMessageActions.dismissErrorMessage(ErrorKeys.BugBashItemError);
    }
}