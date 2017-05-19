import "../../css/BugBashEditor.scss";

import * as React from "react";

import { TextField } from "OfficeFabric/TextField";
import { CommandBar } from "OfficeFabric/CommandBar";
import { DatePicker } from "OfficeFabric/DatePicker";
import { Label } from "OfficeFabric/Label";
import { Dropdown, IDropdownOption, IDropdownProps } from "OfficeFabric/Dropdown";
import { IContextualMenuItem } from "OfficeFabric/components/ContextualMenu/ContextualMenu.Props";
import { autobind } from "OfficeFabric/Utilities";
import { MessageBar, MessageBarType } from "OfficeFabric/MessageBar";
import { Checkbox } from "OfficeFabric/Checkbox";

import { HostNavigationService } from "VSS/SDK/Services/Navigation";
import { WorkItemTemplateReference, WorkItemField, WorkItemType, FieldType } from "TFS/WorkItemTracking/Contracts";
import Utils_String = require("VSS/Utils/String");
import Utils_Date = require("VSS/Utils/Date");

import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { WorkItemFieldStore } from "VSTS_Extension/Stores/WorkItemFieldStore";
import { WorkItemTypeStore } from "VSTS_Extension/Stores/WorkItemTypeStore";
import { WorkItemTemplateStore } from "VSTS_Extension/Stores/WorkItemTemplateStore";
import { BaseStore } from "VSTS_Extension/Stores/BaseStore";
import { Loading } from "VSTS_Extension/Components/Common/Loading";
import { InputError } from "VSTS_Extension/Components/Common/InputError";
import { InfoLabel } from "VSTS_Extension/Components/Common/InfoLabel";
import { RichEditor } from "VSTS_Extension/Components/Common/RichEditor/RichEditor";

import * as Helpers from "../Helpers";
import { StoresHub } from "../Stores/StoresHub";
import { UrlActions } from "../Constants";
import { IBugBash } from "../Interfaces";

export interface IBugBashEditorProps extends IBaseComponentProps {
    id?: string;
}

export interface IBugBashEditorState extends IBaseComponentState {
    originalModel?: IBugBash;
    model?: IBugBash;
    loading?: boolean;
    error?: string;
    disableToolbar?: boolean;
}

export class BugBashEditor extends BaseComponent<IBugBashEditorProps, IBugBashEditorState>  {
    public static getNewModel(): IBugBash {
        return {
            id: "",
            title: "",
            __etag: 0,
            projectId: VSS.getWebContext().project.id,
            workItemType: "",
            itemDescriptionField: "",
            autoAccept: false,
            description: "",
            acceptTemplate: {
                team: VSS.getWebContext().team.id,
                templateId: ""
            }
        };
    }

    protected initializeState(): void {
        let model: IBugBash;
        if (this.props.id) {
            model = StoresHub.bugBashStore.getItem(this.props.id);
        }
        else {
            model = BugBashEditor.getNewModel();
        }

        this.state = {
            model: model,
            originalModel: {...model},
            loading: true,
            error: null
        };
    }

    protected getStoresToLoad(): {new (): BaseStore<any, any, any>}[] {
        return [WorkItemFieldStore, WorkItemTemplateStore, WorkItemTypeStore];
    }

    protected initialize(): void {
        StoresHub.workItemFieldStore.initialize();
        StoresHub.workItemTemplateStore.initialize();
        StoresHub.workItemTypeStore.initialize();
    }

    protected onStoreChanged() {
        this.updateState({
            loading: StoresHub.workItemTypeStore.isLoaded() && StoresHub.workItemFieldStore.isLoaded() && StoresHub.workItemTemplateStore.isLoaded() ? false : true
        });
    }    

    public render(): JSX.Element {
        if (this.state.loading) {
            return <Loading />;
        }

        const model = this.state.model;

        let witItems: IDropdownOption[] = StoresHub.workItemTypeStore.getAll().map((workItemType: WorkItemType, index: number) => {
            return {
                key: workItemType.name,
                index: index + 1,
                text: workItemType.name,
                selected: model.workItemType ? Utils_String.equals(model.workItemType, workItemType.name, true) : false
            }
        });

        let fieldItems: IDropdownOption[] = StoresHub.workItemFieldStore.getAll().filter(f => f.type === FieldType.Html).map((field: WorkItemField, index: number) => {
            return {
                key: field.referenceName,
                index: index + 1,
                text: field.name,
                selected: model.itemDescriptionField ? Utils_String.equals(model.itemDescriptionField, field.referenceName, true) : false
            }
        });

        return (
            <div className="editor-view">
                <CommandBar 
                    className="editor-view-menu"
                    items={this._getCommandBarItems()} 
                    farItems={[{
                        key: "Home", name: "Home", title: "Return to home view", iconProps: {iconName: "Home"}, 
                        onClick: async (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                            let navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
                            navigationService.updateHistoryEntry(UrlActions.ACTION_ALL, null);
                        }
                    }]} />

                { this.state.error && (<MessageBar messageBarType={MessageBarType.error}>{this.state.error}</MessageBar> )}
                
                <div className="editor-view-contents">                    
                    <div className="first-section">                        
                        <TextField 
                            label='Title' 
                            value={model.title} 
                            onChanged={(newValue: string) => this._updateTitle(newValue)} 
                            onGetErrorMessage={this._getTitleError} />

                        <Label>Description</Label>
                        <RichEditor containerId="rich-editor" data={model.description} onChange={(newValue: string) => this._updateDescription(newValue)} />
                    </div>
                    <div className="second-section">
                        <div className="checkbox-container">                            
                            <Checkbox 
                                className="auto-accept"
                                label=""
                                checked={model.autoAccept}
                                onChange={(ev: React.FormEvent<HTMLElement>, isChecked: boolean) => this._updateAutoAccept(isChecked) } />

                            <InfoLabel label="Auto Accept?" info="Auto create work items on creation of a bug bash item" />
                        </div>

                        <DatePicker 
                            label="Start Date" 
                            allowTextInput={true} 
                            isRequired={false} 
                            value={model.startTime}
                            onSelectDate={(newValue: Date) => this._updateStartTime(newValue)} />

                        <DatePicker 
                            label="Finish Date" 
                            allowTextInput={true}
                            isRequired={false} 
                            value={model.endTime} 
                            onSelectDate={(newValue: Date) => this._updateEndTime(newValue)} />

                        { model.startTime && model.endTime && Utils_Date.defaultComparer(model.startTime, model.endTime) >= 0 &&  (<InputError error="Bugbash end time cannot be a date before bugbash start time." />)}

                        <InfoLabel label="Work item type" info="Select a work item type which would be used to create work items for each bug bash item" />
                        <Dropdown 
                            className={!model.workItemType ? "editor-dropdown no-margin" : "editor-dropdown"}
                            onRenderList={this._onRenderCallout} 
                            required={true} 
                            options={witItems}                            
                            onChanged={(option: IDropdownOption) => this._updateWorkItemType(option.key as string)} />

                        { !model.workItemType && (<InputError error="A work item type is required." />) }

                        <InfoLabel label="Description field" info="Select a HTML field that you would want to set while creating a workitem for each bug bash item" />
                        <Dropdown 
                            className={!model.itemDescriptionField ? "editor-dropdown no-margin" : "editor-dropdown"}
                            onRenderList={this._onRenderCallout} 
                            required={true} 
                            options={fieldItems} 
                            onChanged={(option: IDropdownOption) => this._updateDescriptionField(option.key as string)} />

                        { !model.itemDescriptionField && (<InputError error="A description field is required." />) }       

                        <InfoLabel label="Work item template" info="Select a work item template that would be applied during work item creation." />
                        <Dropdown 
                            className="editor-dropdown"
                            onRenderList={this._onRenderCallout} 
                            options={this._getTemplateDropdownOptions(model.acceptTemplate.templateId)} 
                            onChanged={(option: IDropdownOption) => this._updateAcceptTemplate(option.key as string)} />
                    </div>
                </div>
            </div>
        );
    }

    private _getCommandBarItems(): IContextualMenuItem[] {
        const model = this.state.model;

        return [
            {
                key: "save", name: "Save", title: "Save", iconProps: {iconName: "Save"}, disabled: this.state.disableToolbar || !this._isDirty() || !this._isValid(),
                onClick: async (event?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem) => {
                    this.updateState({disableToolbar: true, error: null});

                    if (this._isNew() && this._isDirty() && this._isValid()) {
                        try {
                            let createdModel = await StoresHub.bugBashStore.createItem(model);
                            let navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
                            navigationService.updateHistoryEntry(UrlActions.ACTION_EDIT, { id: createdModel.id }, true, undefined, undefined, true);

                            this.updateState({model: {...createdModel}, originalModel: {...createdModel}, error: null, disableToolbar: false});
                        }
                        catch (e) {
                            this.updateState({error: "We encountered some error while creating the bug bash. Please refresh the page and try again.", disableToolbar: false});
                        }
                    }
                    else if (!this._isNew() && this._isDirty() && this._isValid()) {
                        try {
                            let updatedModel = await StoresHub.bugBashStore.updateItem(model);
                            this.updateState({model: {...updatedModel}, originalModel: {...updatedModel}, error: null, disableToolbar: false});
                        }
                        catch (e) {
                            this.updateState({error: "This bug bash instance has been modified by some one else. Please refresh the page to get the latest version and try updating it again.", disableToolbar: false});
                        }
                    }
                }
            },
            {
                key: "undo", name: "Undo", title: "Undo changes", iconProps: {iconName: "Undo"}, disabled: this.state.disableToolbar || !this._isDirty(),
                onClick: async (event?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem) => {
                    this.updateState({disableToolbar: true});

                    const confirm = await Helpers.confirmAction(true, "Are you sure you want to undo your changes to this instance?");
                    if (confirm) {
                        this.updateState({model: {...this.state.originalModel}, disableToolbar: false, error: null});
                    }
                    else {
                        this.updateState({disableToolbar: false});
                    }
                }
            },
            {
                key: "delete", name: "Delete", title: "Delete", iconProps: {iconName: "Delete"}, disabled: this.state.disableToolbar || this._isNew(),
                onClick: async (event?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem) => {
                    this.updateState({disableToolbar: true});

                    if (!this._isNew()) {
                        const confirm = await Helpers.confirmAction(true, "Are you sure you want to delete this instance?");
                        if (confirm) {
                            try {
                                await StoresHub.bugBashStore.deleteItem(model);
                                let navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
                                navigationService.updateHistoryEntry(UrlActions.ACTION_ALL, null);
                            }
                            catch (e) {
                                this.updateState({error: "We encountered some error while deleting the bug bash. Please refresh the page and try again.", disableToolbar: false});
                            }
                        }
                        else {
                            this.updateState({disableToolbar: false});
                        }             
                    }
                }
            },
            {
                key: "results", name: "Show results", title: "Show results", iconProps: {iconName: "ShowResults"}, disabled: this.state.disableToolbar || this._isNew(),
                onClick: async (event?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem) => {
                    if (!this._isNew()) {
                        const confirm = await Helpers.confirmAction(this._isDirty(), "Are you sure you want to go back to results view? This action will reset your unsaved changes.");
                        if (confirm) {
                            this.updateState({model: {...this.state.originalModel}, error: null});
                            let navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
                            navigationService.updateHistoryEntry(UrlActions.ACTION_VIEW, {id: model.id});
                        }
                    }
                }
            },
        ];
    }

    @autobind
    private _onRenderCallout(props?: IDropdownProps, defaultRender?: (props?: IDropdownProps) => JSX.Element): JSX.Element {
        return (
            <div className="callout-container">
                {defaultRender(props)}
            </div>
        );
    }

    private _getTemplateDropdownOptions(selectedValue: string): IDropdownOption[] {
        let emptyTemplateItem = [
            {   
                key: "", index: 0, text: "<No template>", 
                selected: !selectedValue
            }
        ];
        let filteredTemplates = StoresHub.workItemTemplateStore.getAll().filter((t: WorkItemTemplateReference) => Utils_String.equals(t.workItemTypeName, this.state.model.workItemType));
        return emptyTemplateItem.concat(filteredTemplates.map((template: WorkItemTemplateReference, index: number) => {
            return {
                key: template.id,
                index: index + 1,
                text: template.name,
                selected: selectedValue ? Utils_String.equals(selectedValue, template.id, true) : false
            }
        }));
    }
    
    @autobind
    private _getTitleError(value: string): string | IPromise<string> {
        if (!value) {
            return "Title is required";
        }
        if (value.length > 256) {
            return `The length of the title should less than 256 characters, actual is ${value.length}.`
        }
        return "";
    }

    private _isNew(): boolean {
        return !this.state.model.id;
    }

    private _isDirty(): boolean {        
        return !Utils_String.equals(this.state.model.title, this.state.originalModel.title)
            || !Utils_String.equals(this.state.model.workItemType, this.state.originalModel.workItemType, true)
            || !Utils_String.equals(this.state.model.description, this.state.originalModel.description)
            || !Utils_Date.equals(this.state.model.startTime, this.state.originalModel.startTime)
            || !Utils_Date.equals(this.state.model.endTime, this.state.originalModel.endTime)
            || !Utils_String.equals(this.state.model.itemDescriptionField, this.state.originalModel.itemDescriptionField, true)
            || this.state.model.autoAccept !== this.state.originalModel.autoAccept
            || !Utils_String.equals(this.state.model.acceptTemplate.team, this.state.originalModel.acceptTemplate.team)
            || !Utils_String.equals(this.state.model.acceptTemplate.templateId, this.state.originalModel.acceptTemplate.templateId)
    }

    private _isValid(): boolean {
        return this.state.model.title.trim().length > 0
            && this.state.model.title.length <= 256
            && this.state.model.workItemType.trim().length > 0
            && this.state.model.itemDescriptionField.trim().length > 0
            && (!this.state.model.startTime || !this.state.model.endTime || Utils_Date.defaultComparer(this.state.model.startTime, this.state.model.endTime) < 0);
    }

    private _updateTitle(newTitle: string) {
        let newModel = {...this.state.model};
        newModel.title = newTitle;
        this.updateState({model: newModel});
    }

    private _updateWorkItemType(newType: string) {
        let newModel = {...this.state.model};
        newModel.workItemType = newType;
        newModel.acceptTemplate = {team: VSS.getWebContext().team.id, templateId: ""};  // reset template
        this.updateState({model: newModel});
    }

    private _updateDescription(newDescription: string) {
        let newModel = {...this.state.model};
        newModel.description = newDescription;
        this.updateState({model: newModel});
    }

    private _updateDescriptionField(fieldRefName: string) {
        let newModel = {...this.state.model};
        newModel.itemDescriptionField = fieldRefName;
        this.updateState({model: newModel});
    }

    private _updateStartTime(newStartTime: Date) {
        let newModel = {...this.state.model};
        newModel.startTime = newStartTime;
        this.updateState({model: newModel});
    }

    private _updateEndTime(newEndTime: Date) {        
        let newModel = {...this.state.model};
        newModel.endTime = newEndTime;
        this.updateState({model: newModel});
    }

    private _updateAutoAccept(autoAccept: boolean) {        
        let newModel = {...this.state.model};
        newModel.autoAccept = autoAccept;
        this.updateState({model: newModel});
    }

    private _updateAcceptTemplate(templateId: string) {
        let newModel = {...this.state.model};
        newModel.acceptTemplate = {team: VSS.getWebContext().team.id, templateId: templateId};
        this.updateState({model: newModel});
    }
}