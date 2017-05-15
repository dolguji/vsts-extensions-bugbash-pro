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

import { WorkItem } from "TFS/WorkItemTracking/Contracts";
import Utils_String = require("VSS/Utils/String");
import Utils_Date = require("VSS/Utils/Date");

import { confirmAction, BugBashItemHelpers } from "../Helpers";
import { IBugBashItem, IBugBashItemViewModel, IAcceptedItemViewModel } from "../Interfaces";
import { BugBashItemManager } from "../BugBashItemManager";
import { StoresHub } from "../Stores/StoresHub";
import { RichEditor } from "./RichEditor/RichEditor";

export interface IBugBashItemEditorProps extends IBaseComponentProps {
    viewModel: IBugBashItemViewModel;
    onClickNew: () => void;
    onDelete: (item: IBugBashItem) => void;
    onItemUpdate: (item: IBugBashItem) => void;
    onItemAccept: (item: IBugBashItem, workItem: WorkItem) => void;
    onChange: (changedData: {id: string, title: string, description: string}) => void;
}

export interface IBugBashItemEditorState extends IBaseComponentState {
    loadError?: string;
    viewModel?: IBugBashItemViewModel;
    error?: string;
    disableToolbar?: boolean;
}

export class BugBashItemEditor extends BaseComponent<IBugBashItemEditorProps, IBugBashItemEditorState> {
    protected initializeState() {
        this.state = {
            viewModel: {
                model: BugBashItemHelpers.deepCopy(this.props.viewModel.model),
                originalModel: BugBashItemHelpers.deepCopy(this.props.viewModel.originalModel)
            }
        };
    }

    public componentWillReceiveProps(nextProps: Readonly<IBugBashItemEditorProps>): void {
        if (this.props.viewModel.model.id !== nextProps.viewModel.model.id) {
            this.updateState({
                viewModel: {
                    model: BugBashItemHelpers.deepCopy(nextProps.viewModel.model),
                    originalModel: BugBashItemHelpers.deepCopy(nextProps.viewModel.originalModel)
                }
            });
        }
    }

    private _onChange(newViewModel: IBugBashItemViewModel) {
        this.props.onChange({
            id: newViewModel.model.id,
            title: newViewModel.model.title,
            description: newViewModel.model.description
        });
    }

    public render(): JSX.Element {
        if (this.state.loadError) {
            return <MessageBar messageBarType={MessageBarType.error}>{this.state.loadError}</MessageBar>;
        }
        else if (!this.state.viewModel) {
            return <Loading />;
        }
        else {
            const item = this.state.viewModel;

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
                                let newModel = {...this.state.viewModel};
                                newModel.model.title = newValue;
                                this.updateState({viewModel: newModel});
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
                                let newViewModel = {...this.state.viewModel};
                                newViewModel.model.description = newValue;
                                this.updateState({viewModel: newViewModel});
                                this._onChange(newViewModel);
                            }} />
                    </div>
                </div>
            );
        }        
    }

    private _isNew(): boolean {
        return !this.state.viewModel.model.id;
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
        let menuItems: IContextualMenuItem[] = [
            {
                key: "save", name: "", 
                title: "Save", iconProps: {iconName: "Save"}, 
                disabled: this.state.disableToolbar || !BugBashItemHelpers.isDirty(this.state.viewModel) || !BugBashItemHelpers.isValid(this.state.viewModel.model),
                onClick: async () => {
                    this.updateState({disableToolbar: true, error: null});
                    let updatedModel: IBugBashItem;

                    try {
                        updatedModel = await BugBashItemManager.beginSave(this.state.viewModel.model);
                        this.updateState({error: null, disableToolbar: false,
                            viewModel: BugBashItemHelpers.getItemViewModel(updatedModel)
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

                    const confirm = await confirmAction(BugBashItemHelpers.isDirty(this.state.viewModel), "Refreshing the item will undo your unsaved changes. Are you sure you want to do that?");
                    if (confirm) {
                        const id = this.state.viewModel.model.id;
                        const bugBashId = this.state.viewModel.model.bugBashId;

                        this.updateState({viewModel: null});
                        newModel = await BugBashItemManager.beginGetItem(id, bugBashId);
                        if (newModel) {
                            this.updateState({viewModel: BugBashItemHelpers.getItemViewModel(newModel), error: null, disableToolbar: false});
                            this.props.onItemUpdate(newModel);
                        }
                        else {
                            this.updateState({viewModel: null, error: null, loadError: "This item no longer exist. Please refresh the list and try again."});
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
                disabled: this.state.disableToolbar || !BugBashItemHelpers.isDirty(this.state.viewModel),
                onClick: async () => {
                    this.updateState({disableToolbar: true});

                    const confirm = await confirmAction(true, "Are you sure you want to undo your changes to this item?");
                    if (confirm) {
                        let newViewModel = {...this.state.viewModel};
                        newViewModel.model = BugBashItemHelpers.deepCopy(newViewModel.originalModel);
                        this.updateState({viewModel: newViewModel, disableToolbar: false, error: null});
                        this.props.onItemUpdate(newViewModel.model);
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

                    const confirm = await confirmAction(true, "Are you sure you want to delete this item?");
                    if (confirm) {
                        this.updateState({disableToolbar: false, error: null});
                        try {
                            BugBashItemManager.deleteItems([this.state.viewModel.model]);
                        }
                        catch (e) {
                            
                        }
                        this.props.onDelete(this.state.viewModel.model);
                    }
                    else {
                        this.updateState({disableToolbar: false});
                    }
                }
            }
        ]

        let bugBash = StoresHub.bugBashStore.getItem(this.state.viewModel.model.bugBashId);

        if (!bugBash.autoAccept) {
            menuItems.push({
                key: "Accept", name: "Accept item", title: "Create workitems from selected items", iconProps: {iconName: "Accept"}, 
                disabled: this.state.disableToolbar || BugBashItemHelpers.isDirty(this.state.viewModel) || this._isNew(),
                onClick: async () => {
                    this.updateState({disableToolbar: true});

                    let result: IAcceptedItemViewModel;
                    try {
                        result = await BugBashItemManager.acceptItem(this.state.viewModel.model);                        
                        this.updateState({viewModel: BugBashItemHelpers.getItemViewModel(result.model), error: result.workItem ? null : "Could not create work item. Please refresh the page and try again.", disableToolbar: false});

                        this.props.onItemAccept(result.model, result.workItem);
                    }
                    catch (e) {
                        this.updateState({disableToolbar: false, error: e.message || e});
                    }                        
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

    @autobind
    private _getTitleError(value: string): string | IPromise<string> {
        if (value.length > 256) {
            return `The length of the title should less than 256 characters, actual is ${value.length}.`;
        }
        return "";
    }
}