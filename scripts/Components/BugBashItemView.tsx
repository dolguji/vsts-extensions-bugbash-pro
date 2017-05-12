import "../../css/BugBashItemView.scss";

import * as React from "react";
import { autobind } from "OfficeFabric/Utilities";
import { Label } from "OfficeFabric/Label";
import { TextField } from "OfficeFabric/TextField";
import { PrimaryButton } from "OfficeFabric/Button";
import { MessageBar, MessageBarType } from "OfficeFabric/MessageBar";
import { CommandBar } from "OfficeFabric/CommandBar";
import { IContextualMenuItem } from "OfficeFabric/components/ContextualMenu/ContextualMenu.Props";

import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { BaseStore } from "VSTS_Extension/Stores/BaseStore";
import { Loading } from "VSTS_Extension/Components/Common/Loading";

import { WorkItemFormNavigationService } from "TFS/WorkItemTracking/Services";
import Utils_String = require("VSS/Utils/String");

import { IBugBashItem, IBugBashItemComment } from "../Models";
import Helpers = require("../Helpers");
import { BugBashItemCommentStore } from "../Stores/BugBashItemCommentStore";
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
    model?: IBugBashItem;
    comments?: IBugBashItemComment[];
    loadingComments?: boolean;
    error?: string;
    disableToolbar?: boolean;
}

export class BugBashItemView extends BaseComponent<IBugBashItemViewProps, IBugBashItemViewState> {
    public static getNewModel(bugBashId: string) {
        return {
            id: "",
            bugBashId: bugBashId,
            title: "",
            __etag: 0,
            description: "",
            workItemId: null,
            createdBy: "",
            createdDate: null,
        };
    }

    protected getStoresToLoad(): {new (): BaseStore<any, any, any>}[] {
        return [BugBashItemCommentStore];
    }

    protected initializeState() {
        this._initializeState(this.props);
    }

    protected initialize() {
        this._initialize();
    }   

    protected onStoreChanged() {
        if (this.state.model.id) {
            this.updateState({
                comments: StoresHub.bugBashItemCommentStore.getItems(this.state.model.id),
                loadingComments: !StoresHub.bugBashItemCommentStore.isDataLoaded(this.state.model.id)
            });
        }
    }

    public componentWillReceiveProps(nextProps: Readonly<IBugBashItemViewProps>): void {
        if (!Utils_String.equals(this.props.id, nextProps.id, true)) {
            this._initializeState(nextProps);
            this._initialize();
        }
    }

    private _initialize() {        
        if (this.state.model.id) {
            StoresHub.bugBashItemCommentStore.refreshItems(this.state.model.id);
        }
        else {
            this.updateState({loadingComments: false});
        }
    }

    private _initializeState(props: IBugBashItemViewProps) {
        let model: IBugBashItem;
        if (props.id) {
            model = StoresHub.bugBashItemStore.getItem(props.id);
        }
        else {
            model = BugBashItemView.getNewModel(props.bugBashId);
        }

        this.state = {
            model: model,
            originalModel: {...model},
            comments: [],
            loadingComments: true
        };
    }

    public render(): JSX.Element {
        let model = this.state.model;
        if (!model) {
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
                    <Label className="item-description">Description</Label>
                    <RichEditor containerId="rich-editor" data={model.description} onChange={(newValue: string) => {
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
                        newModel = await StoresHub.bugBashItemStore.refreshItem(this.state.model);
                        this.updateState({model: {...newModel}, originalModel: {...newModel}, error: null, disableToolbar: false});
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
                key: "delete", name: "Delete", title: "Delete", iconProps: {iconName: "Delete"}, disabled: this.state.disableToolbar || this._isNew(),
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