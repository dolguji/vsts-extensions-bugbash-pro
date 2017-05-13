import "../../css/BugBashItemEditor.scss";

import * as React from "react";
import { autobind } from "OfficeFabric/Utilities";
import { Label } from "OfficeFabric/Label";
import { TextField } from "OfficeFabric/TextField";
import { MessageBar, MessageBarType } from "OfficeFabric/MessageBar";
import { CommandBar } from "OfficeFabric/CommandBar";
import { IContextualMenuItem } from "OfficeFabric/components/ContextualMenu/ContextualMenu.Props";

import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { Loading } from "VSTS_Extension/Components/Common/Loading";
import { IdentityView } from "VSTS_Extension/Components/WorkItemControls/IdentityView";

import Utils_String = require("VSS/Utils/String");
import Utils_Date = require("VSS/Utils/Date");

import { IBugBashItem, IComment } from "../Models";
import { StoresHub } from "../Stores/StoresHub";
import { RichEditor } from "./RichEditor/RichEditor";

export interface IBugBashItemEditorProps extends IBaseComponentProps {
    id?: string;
    bugBashId: string;
    onClickNew: () => void;
    onDelete: (item: IBugBashItem) => void;
    onSave: (item: IBugBashItem) => void;
}

export interface IBugBashItemEditorState extends IBaseComponentState {
    originalModel?: IBugBashItem;
    loadError?: string;
    model?: IBugBashItem;
    error?: string;
    disableToolbar?: boolean;
    newComment?: string;
}

export class BugBashItemEditor extends BaseComponent<IBugBashItemEditorProps, IBugBashItemEditorState> {
    public static getNewModel(bugBashId: string): IBugBashItem {
        return {
            id: "",
            bugBashId: bugBashId,
            title: "",
            __etag: 0,
            description: "",
            comments: [],            
            workItemId: null,
            createdBy: "",
            createdDate: null
        };
    }

    protected initializeState() {
        let model = this.props.id ? StoresHub.bugBashItemStore.getItem(this.props.id) : BugBashItemEditor.getNewModel(this.props.bugBashId);

        this.state = {
            model: {...model},
            originalModel: {...model}
        };
    }

    public componentWillReceiveProps(nextProps: Readonly<IBugBashItemEditorProps>): void {
        if (this.props.id !== nextProps.id) {
            let model = nextProps.id ? StoresHub.bugBashItemStore.getItem(nextProps.id) : BugBashItemEditor.getNewModel(nextProps.bugBashId);

            this.updateState({
                model: {...model},
                originalModel: {...model}
            });
        }
    }

    public render(): JSX.Element {
        let model = this.state.model;

        if (this.state.loadError) {

        }
        else if (!model) {
            return <Loading />;
        }

        return (
            <div className="item-editor">
                <CommandBar 
                    className="item-editor-menu"
                    items={this._getToolbarItems()} 
                    farItems={this._getFarMenuItems()}
                />

                { this.state.error && <MessageBar messageBarType={MessageBarType.error}>{this.state.error}</MessageBar>}

                <TextField label="Title" 
                        value={model.title}
                        required={true} 
                        onGetErrorMessage={this._getTitleError}
                        onChanged={(newValue: string) => {
                            let newModel = {...this.state.model};
                            newModel.title = newValue;
                            this.updateState({model: newModel});
                        }} />

                <div className="item-description-container">
                    <Label>Description</Label>
                    <RichEditor 
                        containerId="description-editor" 
                        data={model.description} 
                        editorOptions={{
                            btns: [
                                ['formatting'],
                                ['bold', 'italic'], 
                                ['link'],
                                ['superscript', 'subscript'],
                                ['insertImage'],
                                'btnGrp-lists',
                                ['removeformat'],
                                ['fullscreen']
                            ]
                        }}
                        onChange={(newValue: string) => {
                            let newModel = {...this.state.model};
                            newModel.description = newValue;
                            this.updateState({model: newModel});
                        }} />
                </div>

                <div className="item-discussions-container">
                    <Label>Discussions</Label>
                    <RichEditor 
                        containerId="comment-editor" 
                        data={this.state.newComment || ""}
                        placeholder="Enter a comment"
                        editorOptions={{
                            btns: []
                        }}
                        onChange={(newValue: string) => {                            
                            this.updateState({newComment: newValue});
                        }} />

                    <div className="item-comments">
                        {this._renderComments()}
                    </div>
                </div>
            </div>
        );
    }

    private _renderComments(): React.ReactNode {
        let comments = this.state.model.comments.slice();
        comments.sort((c1, c2) => -1 * Utils_Date.defaultComparer(c1.addedDate, c2.addedDate));

        return comments.map((comment: IComment, index: number) => {
            return (
                <div className="item-comment" key={`${index}`}>
                    <div className="comment-info">
                        <IdentityView identityDistinctName={comment.addedBy} />
                        <div className="comment-added-date">{Utils_Date.friendly(new Date(comment.addedDate as any))}</div>
                    </div>
                    <div className="comment-text" dangerouslySetInnerHTML={{ __html: comment.text }} />
                </div>
            );
        });
    }

    private _isDirty(): boolean {        
        return !Utils_String.equals(this.state.originalModel.title, this.state.model.title)
            || !Utils_String.equals(this.state.originalModel.description, this.state.model.description)
            || (this.state.newComment != null && this.state.newComment.trim() !== "");
    }

    private _isValid(): boolean {
        return this.state.model.title.trim().length > 0 && this.state.model.title.length <= 256;
    }

    private _isNew(): boolean {
        return !this.state.model.id;
    }

    private _getFarMenuItems(): IContextualMenuItem[] {
        return [
            {
                key: "createnew", name: "New", disabled: this._isNew(),
                title: "Create new item", iconProps: {iconName: "Add"}, 
                onClick: () => {
                    this.props.onClickNew();
                }
            }
        ]
    }

    private _getToolbarItems(): IContextualMenuItem[] {
        return [
            {
                key: "save", name: "Save", 
                title: "Save", iconProps: {iconName: "Save"}, 
                disabled: this.state.disableToolbar || !this._isDirty() || !this._isValid(),
                onClick: async () => {
                    this.updateState({disableToolbar: true, error: null});
                    let updatedModel: IBugBashItem;
                    let saveModel = {...this.state.model};
                    if (this.state.newComment != null && this.state.newComment.trim() !== "") {
                        saveModel.comments.push({
                            text: this.state.newComment,
                            addedBy: `${VSS.getWebContext().user.name} <${VSS.getWebContext().user.uniqueName}>`,
                            addedDate: new Date(Date.now())
                        });
                    }

                    if (saveModel.id) {
                        try {                            
                            updatedModel = await StoresHub.bugBashItemStore.updateItem(saveModel);
                            this.updateState({model: {...updatedModel}, originalModel: {...updatedModel}, newComment: null, error: null, disableToolbar: false});
                        }
                        catch (e) {
                            this.updateState({error: "This item has been modified by some one else. Please refresh the item to get the latest version and try updating it again.", disableToolbar: false});
                            return;
                        }
                    }
                    else {
                        try {
                            updatedModel = await StoresHub.bugBashItemStore.createItem(saveModel);
                            this.updateState({model: {...updatedModel}, originalModel: {...updatedModel}, newComment: null, error: null, disableToolbar: false});
                        }
                        catch (e) {
                            this.updateState({error: "We encountered some error while creating the bug bash. Please refresh the page and try again.", disableToolbar: false});
                            return;
                        }
                    }

                    this.props.onSave(updatedModel);
                }
            },
            {
                key: "refresh", name: "Refresh", 
                title: "Refresh", iconProps: {iconName: "Refresh"}, 
                disabled: this.state.disableToolbar || this._isNew(),
                onClick: async () => {
                    this.updateState({disableToolbar: true, error: null});
                    let newModel: IBugBashItem;

                    if (this._isDirty()) {
                        let dialogService: IHostDialogService = await VSS.getService(VSS.ServiceIds.Dialog) as IHostDialogService;
                        try {
                            await dialogService.openMessageDialog("Refreshing the item will undo your unsaved changes. Are you sure you want to do that?", { useBowtieStyle: true });
                        }
                        catch (e) {
                            // user selected "No"" in dialog
                            this.updateState({disableToolbar: false});
                            return;
                        }
                    }

                    try {
                        let toBeRefreshedModel = {...this.state.model};
                        this.updateState({model: null, originalModel: null, newComment: null});
                        newModel = await StoresHub.bugBashItemStore.refreshItem(toBeRefreshedModel);
                        if (newModel) {
                            this.updateState({model: {...newModel}, originalModel: {...newModel}, error: null, disableToolbar: false});
                        }
                        else {
                            this.updateState({model: null, originalModel: null, error: null, loadError: "This item no longer exist. Please refresh the list and try again."});
                        }                        
                    }
                    catch (e) {
                        this.updateState({error: "We encountered some error while creating the bug bash. Please refresh the page and try again.", disableToolbar: false});
                        return;
                    }
                }
            },
            {
                key: "undo", name: "Undo", 
                title: "Undo changes", iconProps: {iconName: "Undo"}, 
                disabled: this.state.disableToolbar || !this._isDirty(),
                onClick: async (event?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem) => {
                    this.updateState({disableToolbar: true});

                    let dialogService: IHostDialogService = await VSS.getService(VSS.ServiceIds.Dialog) as IHostDialogService;
                    try {
                        await dialogService.openMessageDialog("Are you sure you want to undo your changes to this item?", { useBowtieStyle: true });
                    }
                    catch (e) {
                        // user selected "No"" in dialog
                        this.updateState({disableToolbar: false});
                        return;
                    }

                    this.updateState({model: {...this.state.originalModel}, newComment: null, disableToolbar: false, error: null});
                }
            },
            {
                key: "delete", name: "Delete", title: "Delete", iconProps: {iconName: "Delete"}, 
                disabled: this.state.disableToolbar || this._isNew(),
                onClick: async () => {
                    this.updateState({disableToolbar: true});

                    let dialogService: IHostDialogService = await VSS.getService(VSS.ServiceIds.Dialog) as IHostDialogService;
                    try {
                        await dialogService.openMessageDialog("Are you sure you want to delete this item?", { useBowtieStyle: true });
                    }
                    catch (e) {
                        this.updateState({disableToolbar: false});
                        // user selected "No"" in dialog
                        return;
                    }

                    this.updateState({disableToolbar: false, error: null});
                    StoresHub.bugBashItemStore.deleteItems([this.state.model]);
                    this.props.onDelete(this.state.model);
                }
            }
        ]
    }

    @autobind
    private _getTitleError(value: string): string | IPromise<string> {
        if (value.length > 256) {
            return `The length of the title should less than 256 characters, actual is ${value.length}.`;
        }
        return "";
    }
}