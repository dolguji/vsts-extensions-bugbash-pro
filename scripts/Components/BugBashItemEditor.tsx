import "../../css/BugBashItemEditor.scss";

import * as React from "react";
import { autobind } from "OfficeFabric/Utilities";
import { Label } from "OfficeFabric/Label";
import { TextField } from "OfficeFabric/TextField";
import { Checkbox } from "OfficeFabric/Checkbox";
import { CommandBar } from "OfficeFabric/CommandBar";
import { IContextualMenuItem } from "OfficeFabric/components/ContextualMenu/ContextualMenu.Props";

import { MessagePanel, MessageType } from "VSTS_Extension/Components/Common/MessagePanel";
import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { Loading } from "VSTS_Extension/Components/Common/Loading";
import { InputError } from "VSTS_Extension/Components/Common/InputError";
import { ComboBox } from "VSTS_Extension/Components/Common/Combo/Combobox";
import { BaseStore } from "VSTS_Extension/Flux/Stores/BaseStore";
import { IdentityView } from "VSTS_Extension/Components/WorkItemControls/IdentityView";

import { VersionControlChangeType, ItemContentType } from "TFS/VersionControl/Contracts";
import * as GitClient from "TFS/VersionControl/GitRestClient";
import Utils_Date = require("VSS/Utils/Date");
import Utils_Core = require("VSS/Utils/Core");
import { WorkItem } from "TFS/WorkItemTracking/Contracts";
import { WebApiTeam } from "TFS/Core/Contracts";

import { RichEditorComponent } from "./RichEditorComponent";
import { confirmAction, buildGitPush, BugBashItemHelpers } from "../Helpers";
import { IBugBashItem, IBugBashItemComment, IBugBashItemViewModel, IAcceptedItemViewModel } from "../Interfaces";
import { StoresHub } from "../Stores/StoresHub";
import { BugBashItemCommentActions } from "../Actions/BugBashItemCommentActions";

export interface IBugBashItemEditorProps extends IBaseComponentProps {
    viewModel: IBugBashItemViewModel;
    onDelete: (item: IBugBashItem) => void;
    onItemUpdate: (item: IBugBashItem) => void;
    onItemAccept: (item: IBugBashItem, workItem: WorkItem) => void;
    onChange: (changedItem: IBugBashItem, newComment: string) => void;
}

export interface IBugBashItemEditorState extends IBaseComponentState {
    loadError?: string;
    viewModel?: IBugBashItemViewModel;
    error?: string;
    disableToolbar?: boolean;
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

        if (this.props.viewModel.model.id) {
            //StoresHub.bugBashItemCommentStore.ensureComments(this.props.viewModel.model.id);
            BugBashItemCommentActions.initializeComments(this.props.viewModel.model.id);
        } 
    }

    protected getStatesStore(): IBugBashItemEditorState {
        return {
            comments: StoresHub.bugBashItemCommentStore.getItem(this.props.viewModel.model.id)
        };
    }  

    protected initializeState() {
        this.state = {            
            comments: this.props.viewModel.model.id ? null : [],
            viewModel: {
                model: {...this.props.viewModel.model},
                originalModel: {...this.props.viewModel.originalModel},
                newComment: this.props.viewModel.newComment
            }
        };
    }

    public componentWillReceiveProps(nextProps: Readonly<IBugBashItemEditorProps>): void {
        if (this.state.viewModel.model.id !== nextProps.viewModel.model.id) {
            this.updateState({
                comments: nextProps.viewModel.model.id ? null : [],
                viewModel: {
                    newComment: nextProps.viewModel.newComment,
                    model: {...nextProps.viewModel.model},
                    originalModel: {...nextProps.viewModel.originalModel}
                }
            }, () => {
                if (nextProps.viewModel.model.id) {
                    BugBashItemCommentActions.initializeComments(this.props.viewModel.model.id);
                }
            });
        }
    }

    public componentWillUnmount() {
        $(window).off("imagepasted", this._imagePastedHandler);
    } 

    public render(): JSX.Element {
        if (this.state.loadError) {
            return <MessagePanel messageType={MessageType.Error} message={this.state.loadError} />;
        }
        else if (!this.state.viewModel) {
            return <Loading />;
        }
        else if (BugBashItemHelpers.isAccepted(this.state.viewModel.model)) {
            return <MessagePanel 
                messageType={MessageType.Info} 
                message={"This item has been accepted. You can view or edit this item's work item from the \"Accepted items\" tab. Please refresh the list to clear this message."} />;
        }
        else {
            const item = this.state.viewModel;
            const allTeams = StoresHub.teamStore.getAll();
            const team = StoresHub.teamStore.getItem(item.model.teamId);

            return (
                <div className="item-editor" onKeyDown={this._onEditorKeyDown} tabIndex={0}>
                    <CommandBar 
                        className="item-editor-menu"
                        items={this._getToolbarItems()}                         
                    />

                    { this.state.error && <MessagePanel messageType={MessageType.Error} message={this.state.error} />}
                    
                    <TextField label="Title" 
                            value={item.model.title}
                            required={true} 
                            onGetErrorMessage={this._getTitleError}
                            onChanged={(newValue: string) => {
                                let newModel = {...this.state.viewModel};
                                newModel.model.title = newValue;
                                this.updateState({viewModel: newModel});
                                this._onChange(newModel);
                            }} />

                    <div className="item-team-container">
                        <Label required={true}>Team</Label>

                        <ComboBox                             
                            value={team ? team.name : item.model.teamId} 
                            options={{
                                type: "list",
                                mode: "drop",
                                allowEdit: true,
                                source: allTeams.map(t => t.name)
                            }} 
                            onChange={(newTeamName: string) => {
                                let newModel = {...this.state.viewModel};
                                const newTeam = StoresHub.teamStore.getItem(newTeamName);
                                newModel.model.teamId = newTeam ? newTeam.id : newTeamName;
                                this.updateState({viewModel: newModel});
                                this._onChange(newModel);
                            }}/>
                        
                        { this._renderTeamError(item.model.teamId, team) }
                    </div>

                    { item.model.rejected && 
                        <TextField label="Reject reason" 
                                className="reject-reason-input"
                                value={item.model.rejectReason}
                                required={true} 
                                onGetErrorMessage={this._getRejectReasonError}
                                onChanged={(newValue: string) => {
                                    let newModel = {...this.state.viewModel};
                                    newModel.model.rejectReason = newValue;
                                    this.updateState({viewModel: newModel});
                                    this._onChange(newModel);
                                }} />
                    }

                    <div className="item-description-container">
                        <Label>Description</Label>
                        <RichEditorComponent 
                            containerId="description-editor" 
                            data={item.model.description} 
                            editorOptions={{
                                svgPath: `${VSS.getExtensionContext().baseUri}/css/libs/icons.svg`,
                                btns: [
                                    ['formatting'],
                                    ['bold', 'italic'], 
                                    ['link'],
                                    ['insertImage'],
                                    'btnGrp-lists',
                                    ['removeformat'],
                                    ['fullscreen']
                                ]
                            }}
                            onChange={(newValue: string) => {
                                let newViewModel = {...this.state.viewModel};
                                newViewModel.model.description = newValue;
                                this.updateState({viewModel: newViewModel});
                                this._onChange(newViewModel);
                            }} />
                    </div>

                    <div className="item-comments-editor-container">
                        <Label>Discussion</Label>
                        <RichEditorComponent 
                            containerId="comment-editor" 
                            data={this.state.viewModel.newComment || ""} 
                            editorOptions={{
                                svgPath: `${VSS.getExtensionContext().baseUri}/css/libs/icons.svg`,
                                btns: []
                            }}
                            onChange={(newComment: string) => {     
                                let newViewModel = {...this.state.viewModel};   
                                newViewModel.newComment = newComment;                        
                                this.updateState({viewModel: newViewModel});
                                this._onChange(newViewModel);
                            }} />
                    </div>

                    <div className="item-comments-container">
                        {this._renderComments()}
                    </div>
                </div>
            );
        }        
    }

    private _onChange(newViewModel: IBugBashItemViewModel) {
        this.props.onChange(newViewModel.model, newViewModel.newComment);
    }    

    private _renderComments(): React.ReactNode {
        if (!this._isNew()) {
            let comments = StoresHub.bugBashItemCommentStore.getItem(this.state.viewModel.model.id)
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
            this._saveItem();
        }
    }

    private _isNew(): boolean {
        return !this.state.viewModel.model.id;
    }

    private async _saveItem() {
        const newComment = this.state.viewModel.newComment;
        const isNew = this._isNew();

        if (!this.state.disableToolbar && BugBashItemHelpers.isDirty(this.state.viewModel) && BugBashItemHelpers.isValid(this.state.viewModel.model)) {
            const bugBash = StoresHub.bugBashStore.getItem(this.state.viewModel.model.bugBashId);
            if (isNew && bugBash.autoAccept && !BugBashItemHelpers.isAccepted(this.state.viewModel.model)) {
                this._acceptItem();
            }
            else {
                this.updateState({disableToolbar: true, error: null});
                let updatedModel: IBugBashItem;

                try {
                    // updatedModel = await BugBashItemManager.saveItem(this.state.viewModel.model);

                    // if (newComment && newComment.trim() !== "") {
                    //     // add comment
                    //     StoresHub.bugBashItemCommentStore.addCommentItem(updatedModel.id, newComment);
                    // }
                    // else if (isNew) {
                    //     StoresHub.bugBashItemCommentStore.ensureComments(updatedModel.id);
                    // }
                    
                    // this.updateState({error: null, disableToolbar: false,
                    //     viewModel: BugBashItemHelpers.getItemViewModel(updatedModel)
                    // });
                    // this.props.onItemUpdate(updatedModel);
                }
                catch (e) {
                    this.updateState({error: "This item has been modified by some one else. Please refresh the item to get the latest version and try updating it again.", disableToolbar: false});
                } 
            }
        }
    }

    private async _acceptItem() {
        this.updateState({disableToolbar: true});
        
        try {
            //const result = await BugBashItemManager.acceptItem(this.state.viewModel.model);                        
            //this.updateState({viewModel: BugBashItemHelpers.getItemViewModel(result.model), error: result.workItem ? null : result.error || "Could not create work item. Please refresh the page and try again.", disableToolbar: false});

            //this.props.onItemAccept(result.model, result.workItem);
        }
        catch (e) {
            this.updateState({disableToolbar: false, error: e.message || e});
        }
    }

    private _getToolbarItems(): IContextualMenuItem[] {
        let menuItems: IContextualMenuItem[] = [
            {
                key: "save", name: "", 
                title: "Save", iconProps: {iconName: "Save"}, 
                disabled: this.state.disableToolbar || !BugBashItemHelpers.isDirty(this.state.viewModel) || !BugBashItemHelpers.isValid(this.state.viewModel.model),
                onClick: async () => {
                    this._saveItem();
                }
            },
            {
                key: "refresh", name: "", 
                title: "Refresh", iconProps: {iconName: "Refresh"}, 
                disabled: this.state.disableToolbar || this._isNew(),
                onClick: async () => {
                    this.updateState({disableToolbar: true, error: null});
                    let newModel: IBugBashItem;

                    const confirm = await confirmAction(BugBashItemHelpers.isDirty(this.state.viewModel), "Refreshing the item will undo your unsaved changes. Are you sure you want to do that?");
                    if (confirm) {
                        const id = this.state.viewModel.model.id;
                        const bugBashId = this.state.viewModel.model.bugBashId;

                        // this.updateState({viewModel: null});                        
                        // newModel = await BugBashItemManager.getItem(id, bugBashId);
                        // if (newModel) {
                        //     this.updateState({viewModel: BugBashItemHelpers.getItemViewModel(newModel), error: null, disableToolbar: false});
                        //     this.props.onItemUpdate(newModel);

                        //     StoresHub.bugBashItemCommentStore.refreshComments(id);
                        // }
                        // else {
                        //     this.updateState({viewModel: null, error: null, loadError: "This item no longer exist. Please refresh the list and try again."});
                        // }
                    }
                    else {
                        this.updateState({disableToolbar: false});                    
                    }                
                }
            },
            {
                key: "undo", name: "", 
                title: "Undo changes", iconProps: {iconName: "Undo"}, 
                disabled: this.state.disableToolbar || !BugBashItemHelpers.isDirty(this.state.viewModel),
                onClick: async () => {
                    this.updateState({disableToolbar: true});

                    const confirm = await confirmAction(true, "Are you sure you want to undo your changes to this item?");
                    if (confirm) {
                        let newViewModel = {...this.state.viewModel};
                        newViewModel.model = {...newViewModel.originalModel};
                        newViewModel.newComment = "";
                        this.updateState({viewModel: newViewModel, disableToolbar: false, error: null});
                        this.props.onItemUpdate(newViewModel.model);
                    }
                    else {
                        this.updateState({disableToolbar: false});
                    }
                }
            }
        ]

        let bugBash = StoresHub.bugBashStore.getItem(this.state.viewModel.model.bugBashId);

        if (!bugBash.autoAccept) {
            const isMenuDisabled = this.state.disableToolbar || BugBashItemHelpers.isDirty(this.state.viewModel) || this._isNew();
            menuItems.push(
                {
                    key: "Accept", name: "Accept", title: "Create workitems from selected items", iconProps: {iconName: "Accept"}, className: !isMenuDisabled ? "acceptItemButton" : "",
                    disabled: isMenuDisabled,
                    onClick: () => {
                        this._acceptItem();
                    }
                },
                {
                    key: "Reject",
                    onRender:(item) => {
                        return <Checkbox
                                    disabled={this.state.disableToolbar || this._isNew()} 
                                    className="reject-menu-item-checkbox"
                                    label="Reject"
                                    checked={this.state.viewModel.model.rejected === true}
                                    onChange={(ev: React.FormEvent<HTMLElement>, isChecked: boolean) => {
                                        let newViewModel = {...this.state.viewModel};
                                        newViewModel.model.rejected = !newViewModel.model.rejected;
                                        newViewModel.model.rejectedBy = `${VSS.getWebContext().user.name} <${VSS.getWebContext().user.uniqueName}>`;
                                        newViewModel.model.rejectReason = "";
                                        this.updateState({viewModel: newViewModel});
                                        this._onChange(newViewModel);
                                    }} />;
                    }
                });
        }
        else {
            menuItems.push({
                key: "Accept", name: "Auto accept on", className: "auto-accept-menuitem",
                title: "Auto accept is turned on for this bug bash. A work item would be created as soon as a bug bash item is created", 
                iconProps: {iconName: "SkypeCircleCheck"}
            });
        }

        return menuItems;
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

    private async _onImagePaste(event, args) {
        const data = args.data;
        const callback = args.callback;

        const settings = StoresHub.settingsStore.getAll();
        if (settings && settings.gitMediaRepo) {
            const dataStartIndex = data.indexOf(",") + 1;
            const metaPart = data.substring(5, dataStartIndex - 1);
            const dataPart = data.substring(dataStartIndex);

            const extension = metaPart.split(";")[0].split("/").pop();
            const fileName = `pastedImage_${Date.now().toString()}.${extension}`;
            const gitPath = `BugBash_${StoresHub.bugBashStore.getItem(this.state.viewModel.model.bugBashId).title.replace(" ", "_")}/pastedImages/${fileName}`;
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