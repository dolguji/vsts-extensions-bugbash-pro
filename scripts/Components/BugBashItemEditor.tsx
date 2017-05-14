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

import * as Helpers from "../Helpers";
import { IBugBashItem, IBugBashItemModel, IComment } from "../Models";
import { BugBashItemManager } from "../BugBashItemManager";
import { StoresHub } from "../Stores/StoresHub";
import { RichEditor } from "./RichEditor/RichEditor";

export interface IBugBashItemEditorProps extends IBaseComponentProps {
    itemModel: IBugBashItemModel;
    onClickNew: () => void;
    onDelete: (item: IBugBashItem) => void;
    onItemUpdate: (item: IBugBashItem) => void;
    onChange: (data: {id: string, title: string, description: string, newComment: string}) => void;
}

export interface IBugBashItemEditorState extends IBaseComponentState {
    loadError?: string;
    itemModel?: IBugBashItemModel;
    error?: string;
    disableToolbar?: boolean;
}

export class BugBashItemEditor extends BaseComponent<IBugBashItemEditorProps, IBugBashItemEditorState> {
    protected initializeState() {
        this.state = {
            itemModel: {
                model: BugBashItemManager.deepCopy(this.props.itemModel.model),
                originalModel: BugBashItemManager.deepCopy(this.props.itemModel.originalModel),
                newComment: this.props.itemModel.newComment
            }
        };
    }

    public componentWillReceiveProps(nextProps: Readonly<IBugBashItemEditorProps>): void {
        if (this.props.itemModel.model.id !== nextProps.itemModel.model.id) {
            this.updateState({
                itemModel: {
                    model: BugBashItemManager.deepCopy(nextProps.itemModel.model),
                    originalModel: BugBashItemManager.deepCopy(nextProps.itemModel.originalModel),
                    newComment: nextProps.itemModel.newComment
                }
            });
        }
    }

    private _onChange(newModel: IBugBashItemModel) {
        this.props.onChange({
            id: newModel.model.id,
            title: newModel.model.title,
            description: newModel.model.description,
            newComment: newModel.newComment
        });
    }

    public render(): JSX.Element {
        if (this.state.loadError) {
            return <MessageBar messageBarType={MessageBarType.error}>{this.state.loadError}</MessageBar>;
        }
        else if (!this.state.itemModel) {
            return <Loading />;
        }
        else {
            const item = this.state.itemModel;

            return (
                <div className="item-editor">
                    <CommandBar 
                        className="item-editor-menu"
                        items={this._getToolbarItems()} 
                        farItems={this._getFarMenuItems()}
                    />

                    { this.state.error && <MessageBar messageBarType={MessageBarType.error}>{this.state.error}</MessageBar>}

                    <TextField label="Title" 
                            value={item.model.title}
                            required={true} 
                            onGetErrorMessage={this._getTitleError}
                            onChanged={(newValue: string) => {
                                let newModel = {...this.state.itemModel};
                                newModel.model.title = newValue;
                                this.updateState({itemModel: newModel});
                                this._onChange(newModel);
                            }} />

                    <div className="item-description-container">
                        <Label>Description</Label>
                        <RichEditor 
                            containerId="description-editor" 
                            data={item.model.description} 
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
                                let newModel = {...this.state.itemModel};
                                newModel.model.description = newValue;
                                this.updateState({itemModel: newModel});
                                this._onChange(newModel);
                            }} />
                    </div>

                    <div className="item-discussions-container">
                        <Label>Discussions</Label>
                        <RichEditor 
                            containerId="comment-editor" 
                            data={item.newComment || ""}
                            placeholder="Enter a comment"
                            editorOptions={{
                                btns: []
                            }}
                            onChange={(newValue: string) => {                            
                                let newModel = {...this.state.itemModel};
                                newModel.newComment = newValue;
                                this.updateState({itemModel: newModel});
                                this._onChange(newModel);
                            }} />

                        <div className="item-comments">
                            {this._renderComments()}
                        </div>
                    </div>
                </div>
            );
        }        
    }

    private _renderComments(): React.ReactNode {
        let comments = this.state.itemModel.model.comments.slice();
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

    private _isNew(): boolean {
        return !this.state.itemModel.model.id;
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
                key: "save", name: "", 
                title: "Save", iconProps: {iconName: "Save"}, 
                disabled: this.state.disableToolbar || !BugBashItemManager.isDirty(this.state.itemModel) || !BugBashItemManager.isValid(this.state.itemModel.model),
                onClick: async () => {
                    this.updateState({disableToolbar: true, error: null});
                    let updatedModel: IBugBashItem;

                    try {
                        updatedModel = await BugBashItemManager.beginSave(this.state.itemModel.model, this.state.itemModel.newComment);
                        this.updateState({error: null, disableToolbar: false,
                            itemModel: BugBashItemManager.getItemModel(updatedModel)
                        });
                        this.props.onItemUpdate(updatedModel);
                    }
                    catch (e) {
                        this.updateState({error: "This item has been modified by some one else. Please refresh the item to get the latest version and try updating it again.", disableToolbar: false});
                    }                   
                    
                }
            },
            {
                key: "refresh", name: "", 
                title: "Refresh", iconProps: {iconName: "Refresh"}, 
                disabled: this.state.disableToolbar || this._isNew(),
                onClick: async () => {
                    this.updateState({disableToolbar: true, error: null});
                    let newModel: IBugBashItem;

                    const confirm = await Helpers.confirmAction(BugBashItemManager.isDirty(this.state.itemModel), "Refreshing the item will undo your unsaved changes. Are you sure you want to do that?");
                    if (confirm) {
                        const id = this.state.itemModel.model.id;
                        const bugBashId = this.state.itemModel.model.bugBashId;

                        this.updateState({itemModel: null});
                        newModel = await BugBashItemManager.beginGetItem(id, bugBashId);
                        if (newModel) {
                            this.updateState({itemModel: BugBashItemManager.getItemModel(newModel), error: null, disableToolbar: false});
                            this.props.onItemUpdate(newModel);
                        }
                        else {
                            this.updateState({itemModel: null, error: null, loadError: "This item no longer exist. Please refresh the list and try again."});
                        }
                    }
                    else {
                        this.updateState({disableToolbar: false});                    
                    }                
                }
            },
            {
                key: "undo", name: "", 
                title: "Undo changes", iconProps: {iconName: "Undo"}, 
                disabled: this.state.disableToolbar || !BugBashItemManager.isDirty(this.state.itemModel),
                onClick: async (event?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem) => {
                    this.updateState({disableToolbar: true});

                    const confirm = await Helpers.confirmAction(true, "Are you sure you want to undo your changes to this item?");
                    if (confirm) {
                        let newItemModel = {...this.state.itemModel};
                        newItemModel.model = BugBashItemManager.deepCopy(newItemModel.originalModel);
                        newItemModel.newComment = "";
                        this.updateState({itemModel: newItemModel, disableToolbar: false, error: null});
                        this.props.onItemUpdate(newItemModel.model);
                    }
                    else {
                        this.updateState({disableToolbar: false});
                    }
                }
            },
            {
                key: "delete", name: "", title: "Delete", iconProps: {iconName: "Delete"}, 
                disabled: this.state.disableToolbar || this._isNew(),
                onClick: async () => {
                    this.updateState({disableToolbar: true});

                    const confirm = await Helpers.confirmAction(true, "Are you sure you want to delete this item?");
                    if (confirm) {
                        this.updateState({disableToolbar: false, error: null});
                        try {
                            BugBashItemManager.deleteItems([this.state.itemModel.model]);
                        }
                        catch (e) {
                            
                        }
                        this.props.onDelete(this.state.itemModel.model);
                    }
                    else {
                        this.updateState({disableToolbar: false});
                    }
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