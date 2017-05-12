import "../../css/BugBashItemView.scss";

import * as React from "react";
import { autobind } from "OfficeFabric/Utilities";
import { Label } from "OfficeFabric/Label";
import { TextField } from "OfficeFabric/TextField";
import { MessageBar, MessageBarType } from "OfficeFabric/MessageBar";
import { CommandBar } from "OfficeFabric/CommandBar";
import { IContextualMenuItem } from "OfficeFabric/components/ContextualMenu/ContextualMenu.Props";

import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { Loading } from "VSTS_Extension/Components/Common/Loading";

import Utils_String = require("VSS/Utils/String");

import { IBugBashItem } from "../Models";
import { StoresHub } from "../Stores/StoresHub";
import { RichEditor } from "./RichEditor/RichEditor";

export interface IBugBashItemViewProps extends IBaseComponentProps {
    id?: string;
    bugBashId: string;
    onClickNew: () => void;
    onDelete: (item: IBugBashItem) => void;
    onSave: (item: IBugBashItem) => void;
}

export interface IBugBashItemViewState extends IBaseComponentState {
    originalModel?: IBugBashItem;
    loadError?: string;
    model?: IBugBashItem;
    error?: string;
    disableToolbar?: boolean;
}

export class BugBashItemView extends BaseComponent<IBugBashItemViewProps, IBugBashItemViewState> {
    public static getNewModel(bugBashId: string): IBugBashItem {
        return {
            id: "",
            bugBashId: bugBashId,
            title: "",
            __etag: 0,
            description: "",
            workItemId: null,
            createdBy: "",
            createdDate: null
        };
    }

    protected initializeState() {
        let model = this.props.id ? StoresHub.bugBashItemStore.getItem(this.props.id) : BugBashItemView.getNewModel(this.props.bugBashId);

        this.state = {
            model: {...model},
            originalModel: {...model}
        };
    }

    public componentWillReceiveProps(nextProps: Readonly<IBugBashItemViewProps>): void {
        if (this.props.id !== nextProps.id) {
            let model = nextProps.id ? StoresHub.bugBashItemStore.getItem(nextProps.id) : BugBashItemView.getNewModel(nextProps.bugBashId);

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
            <div className="item-view">
                <CommandBar 
                    className="item-view-menu"
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

                <div>
                    <Label>Description</Label>
                    <RichEditor 
                        containerId="rich-editor" 
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
            </div>
        );
    }

    private _isDirty(): boolean {        
        return !Utils_String.equals(this.state.originalModel.title, this.state.model.title)
            || !Utils_String.equals(this.state.originalModel.description, this.state.model.description);
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

                    if (this.state.model.id) {
                        try {
                            updatedModel = await StoresHub.bugBashItemStore.updateItem(this.state.model);
                            this.updateState({model: {...updatedModel}, originalModel: {...updatedModel}, error: null, disableToolbar: false});
                        }
                        catch (e) {
                            this.updateState({error: "This item has been modified by some one else. Please refresh the item to get the latest version and try updating it again.", disableToolbar: false});
                            return;
                        }
                    }
                    else {
                        try {
                            updatedModel = await StoresHub.bugBashItemStore.createItem(this.state.model);
                            this.updateState({model: {...updatedModel}, originalModel: {...updatedModel}, error: null, disableToolbar: false});
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
                        this.updateState({model: null, originalModel: null});
                        newModel = await StoresHub.bugBashItemStore.refreshItem(this.state.model);
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

                    this.updateState({model: {...this.state.originalModel}, disableToolbar: false, error: null});
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