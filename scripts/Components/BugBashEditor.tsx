import "../../css/BugBashEditor.scss";

import * as React from "react";

import { TextField } from "OfficeFabric/TextField";
import { CommandBar } from "OfficeFabric/CommandBar";
import { DatePicker } from "OfficeFabric/DatePicker";
import { Label } from "OfficeFabric/Label";
import { Dropdown, IDropdownOption, IDropdownProps } from "OfficeFabric/Dropdown";
import { IContextualMenuItem } from "OfficeFabric/components/ContextualMenu/ContextualMenu.Props";
import { autobind } from "OfficeFabric/Utilities";
import { Checkbox } from "OfficeFabric/Checkbox";

import { HostNavigationService } from "VSS/SDK/Services/Navigation";
import { WorkItemTemplateReference, WorkItemField, WorkItemType, FieldType } from "TFS/WorkItemTracking/Contracts";
import Utils_String = require("VSS/Utils/String");
import Utils_Date = require("VSS/Utils/Date");
import Utils_Core = require("VSS/Utils/Core");

import { MessagePanel, MessageType } from "VSTS_Extension/Components/Common/MessagePanel";
import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { InputError } from "VSTS_Extension/Components/Common/InputError";
import { InfoLabel } from "VSTS_Extension/Components/Common/InfoLabel";

import { confirmAction } from "../Helpers";
import { StoresHub } from "../Stores/StoresHub";
import { UrlActions } from "../Constants";
import { IBugBash } from "../Interfaces";
import { RichEditorComponent } from "./RichEditorComponent";
import { BugBashActions } from "../Actions/BugBashActions";

export interface IBugBashEditorProps extends IBaseComponentProps {
    bugBash: IBugBash;
}

export interface IBugBashEditorState extends IBaseComponentState {
    originalBugBash: IBugBash;
    bugBash: IBugBash;
    error?: string;
}

export class BugBashEditor extends BaseComponent<IBugBashEditorProps, IBugBashEditorState>  {
    private _imagePastedHandler: (event, data) => void;    

    constructor(props: IBugBashEditorProps, context?: any) {
        super(props, context);
        this._imagePastedHandler = Utils_Core.delegate(this, this._onImagePaste);
    }

    protected initializeState(): void {
        this.state = {
            bugBash: {...this.props.bugBash},
            originalBugBash: {...this.props.bugBash},
            error: null
        };
    }

    public componentDidMount(): void {
        super.componentDidMount();

        $(window).off("imagepasted", this._imagePastedHandler);
        $(window).on("imagepasted", this._imagePastedHandler);
    }     
    
    public componentWillUnmount() {
        super.componentWillUnmount();
        $(window).off("imagepasted", this._imagePastedHandler);
    }      

    public componentWillReceiveProps(nextProps: Readonly<IBugBashEditorProps>): void {
        if (nextProps.bugBash.id !== this.props.bugBash.id || nextProps.bugBash.__etag !== this.props.bugBash.__etag) {
            this.updateState({
                bugBash: {...nextProps.bugBash},
                originalBugBash: {...nextProps.bugBash},
                error: null
            });
        }
        else {
            this.updateState({
                error: null
            } as IBugBashEditorState);
        }
    }

    public render(): JSX.Element {        
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

                { this.state.error && <MessagePanel messageType={MessageType.Error} message={this.state.error} /> }
                { this._renderEditor() }                
            </div>
        );
    }

    private _renderEditor(): JSX.Element {
        const bugBash = this.state.bugBash;

        const workItemTypes: IDropdownOption[] = StoresHub.workItemTypeStore.getAll().map((workItemType: WorkItemType, index: number) => {
            return {
                key: workItemType.name,
                index: index + 1,
                text: workItemType.name,
                selected: bugBash.workItemType ? Utils_String.equals(bugBash.workItemType, workItemType.name, true) : false
            }
        });

        const htmlFields: IDropdownOption[] = StoresHub.workItemFieldStore.getAll().filter(f => f.type === FieldType.Html).map((field: WorkItemField, index: number) => {
            return {
                key: field.referenceName,
                index: index + 1,
                text: field.name,
                selected: bugBash.itemDescriptionField ? Utils_String.equals(bugBash.itemDescriptionField, field.referenceName, true) : false
            }
        });
        
        return <div className="editor-view-contents" onKeyDown={this._onEditorKeyDown} tabIndex={0}>
            <div className="first-section">                        
                <TextField 
                    label='Title' 
                    value={bugBash.title} 
                    onChanged={(newValue: string) => this._updateTitle(newValue)} 
                    onGetErrorMessage={this._getTitleError} />

                <Label>Description</Label>
                <RichEditorComponent 
                    containerId="bugbash-description-editor" 
                    data={bugBash.description} 
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
                    onChange={(newValue: string) => this._updateDescription(newValue)} />
            </div>
            <div className="second-section">
                <div className="checkbox-container">                            
                    <Checkbox 
                        className="auto-accept"
                        label=""
                        checked={bugBash.autoAccept}
                        onChange={(ev: React.FormEvent<HTMLElement>, isChecked: boolean) => this._updateAutoAccept(isChecked) } />

                    <InfoLabel label="Auto Accept?" info="Auto create work items on creation of a bug bash item" />
                </div>

                <DatePicker 
                    label="Start Date" 
                    allowTextInput={true} 
                    isRequired={false} 
                    value={bugBash.startTime}
                    onSelectDate={(newValue: Date) => this._updateStartTime(newValue)} />

                <DatePicker 
                    label="Finish Date" 
                    allowTextInput={true}
                    isRequired={false} 
                    value={bugBash.endTime} 
                    onSelectDate={(newValue: Date) => this._updateEndTime(newValue)} />

                { bugBash.startTime && bugBash.endTime && Utils_Date.defaultComparer(bugBash.startTime, bugBash.endTime) >= 0 &&  (<InputError error="Bugbash end time cannot be a date before bugbash start time." />)}

                <InfoLabel label="Work item type" info="Select a work item type which would be used to create work items for each bug bash item" />
                <Dropdown 
                    className={!bugBash.workItemType ? "editor-dropdown no-margin" : "editor-dropdown"}
                    onRenderList={this._onRenderCallout} 
                    required={true} 
                    options={workItemTypes}                            
                    onChanged={(option: IDropdownOption) => this._updateWorkItemType(option.key as string)} />

                { !bugBash.workItemType && (<InputError error="A work item type is required." />) }

                <InfoLabel label="Description field" info="Select a HTML field that you would want to set while creating a workitem for each bug bash item" />
                <Dropdown 
                    className={!bugBash.itemDescriptionField ? "editor-dropdown no-margin" : "editor-dropdown"}
                    onRenderList={this._onRenderCallout} 
                    required={true} 
                    options={htmlFields} 
                    onChanged={(option: IDropdownOption) => this._updateDescriptionField(option.key as string)} />

                { !bugBash.itemDescriptionField && (<InputError error="A description field is required." />) }       

                <InfoLabel label="Work item template" info="Select a work item template that would be applied during work item creation." />
                <Dropdown 
                    className="editor-dropdown"
                    onRenderList={this._onRenderCallout} 
                    options={this._getTemplateDropdownOptions(bugBash.acceptTemplate.templateId)} 
                    onChanged={(option: IDropdownOption) => this._updateAcceptTemplate(option.key as string)} />
            </div>
        </div>;
    }

    private async _save() {
        if (this._isNew() && this._isDirty() && this._isValid()) {
            try {
                const createdBugBash = await BugBashActions.createBugBash(this.state.bugBash);
                let navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
                navigationService.updateHistoryEntry(UrlActions.ACTION_EDIT, { id: createdBugBash.id }, true);
            }
            catch (e) {
                this.updateState({error: e} as IBugBashEditorState);
            }
        }
        else if (!this._isNew() && this._isDirty() && this._isValid()) {
            try {
                await BugBashActions.updateBugBash(this.state.bugBash);
            }
            catch (e) {
                this.updateState({error: e} as IBugBashEditorState);
            }
        }
    }

    private _getCommandBarItems(): IContextualMenuItem[] {
        const bugBash = this.state.bugBash;

        return [
            {
                key: "save", name: "Save", title: "Save", iconProps: {iconName: "Save"}, disabled: !this._isDirty() || !this._isValid(),
                onClick: (event?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem) => {
                    this._save();
                }
            },
            {
                key: "undo", name: "Undo", title: "Undo changes", iconProps: {iconName: "Undo"}, disabled: this._isNew() || !this._isDirty(),
                onClick: async (event?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem) => {
                    const confirm = await confirmAction(true, "Are you sure you want to undo your changes to this instance?");
                    if (confirm) {
                        this.updateState({bugBash: {...this.state.originalBugBash}, error: null} as IBugBashEditorState);
                    }
                }
            },
            {
                key: "refresh", name: "Refresh", title: "Refresh", iconProps: {iconName: "Refresh"}, disabled: this._isNew(),
                onClick: async (event?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem) => {
                    const confirm = await confirmAction(this._isDirty(), "Refreshing the item will undo your unsaved changes. Are you sure you want to do that?");
                    if (confirm) {
                        try {
                            await BugBashActions.refreshBugBash(this.state.bugBash.id);
                        }
                        catch (e) {
                            this.updateState({error: e} as IBugBashEditorState);
                        }                        
                    }
                }
            },
            {
                key: "delete", name: "Delete", title: "Delete", iconProps: {iconName: "Cancel"}, disabled: this._isNew(),
                onClick: async (event?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem) => {
                    if (!this._isNew()) {
                        const confirm = await confirmAction(true, "Are you sure you want to delete this instance?");
                        if (confirm) {
                            await BugBashActions.deleteBugBash(bugBash.id);
                            let navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
                            navigationService.updateHistoryEntry(UrlActions.ACTION_ALL, null);
                        }       
                    }
                }
            },
            {
                key: "results", name: "Show results", title: "Show results", iconProps: {iconName: "ShowResults"}, disabled: this._isNew(),
                onClick: async (event?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem) => {
                    if (!this._isNew()) {
                        const confirm = await confirmAction(this._isDirty(), "Are you sure you want to go back to results view? This action will reset your unsaved changes.");
                        if (confirm) {
                            let navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
                            navigationService.updateHistoryEntry(UrlActions.ACTION_VIEW, {id: bugBash.id});
                        }
                    }
                }
            },
        ];
    }

    private _onImagePaste(event, args) {
        args.callback(null);
    }

    @autobind
    private _onEditorKeyDown(e: React.KeyboardEvent<any>) {
        if (e.ctrlKey && e.keyCode === 83) {
            e.preventDefault();
            this._save();
        }
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
        let filteredTemplates = StoresHub.workItemTemplateStore.getAll().filter((t: WorkItemTemplateReference) => Utils_String.equals(t.workItemTypeName, this.state.bugBash.workItemType));
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
        return !this.state.bugBash.id;
    }

    private _isDirty(): boolean {        
        return !Utils_String.equals(this.state.bugBash.title, this.state.originalBugBash.title)
            || !Utils_String.equals(this.state.bugBash.workItemType, this.state.originalBugBash.workItemType, true)
            || !Utils_String.equals(this.state.bugBash.description, this.state.originalBugBash.description)
            || !Utils_Date.equals(this.state.bugBash.startTime, this.state.originalBugBash.startTime)
            || !Utils_Date.equals(this.state.bugBash.endTime, this.state.originalBugBash.endTime)
            || !Utils_String.equals(this.state.bugBash.itemDescriptionField, this.state.originalBugBash.itemDescriptionField, true)
            || this.state.bugBash.autoAccept !== this.state.originalBugBash.autoAccept
            || !Utils_String.equals(this.state.bugBash.acceptTemplate.team, this.state.originalBugBash.acceptTemplate.team)
            || !Utils_String.equals(this.state.bugBash.acceptTemplate.templateId, this.state.originalBugBash.acceptTemplate.templateId)
    }

    private _isValid(): boolean {
        return this.state.bugBash.title.trim().length > 0
            && this.state.bugBash.title.length <= 256
            && this.state.bugBash.workItemType.trim().length > 0
            && this.state.bugBash.itemDescriptionField.trim().length > 0
            && (!this.state.bugBash.startTime || !this.state.bugBash.endTime || Utils_Date.defaultComparer(this.state.bugBash.startTime, this.state.bugBash.endTime) < 0);
    }

    private _updateTitle(newTitle: string) {
        let bugBash = {...this.state.bugBash};
        bugBash.title = newTitle;
        this.updateState({bugBash: bugBash} as IBugBashEditorState);
    }

    private _updateWorkItemType(newType: string) {
        let bugBash = {...this.state.bugBash};
        bugBash.workItemType = newType;
        bugBash.acceptTemplate = {team: VSS.getWebContext().team.id, templateId: ""};  // reset template
        this.updateState({bugBash: bugBash} as IBugBashEditorState);
    }

    private _updateDescription(newDescription: string) {
        let bugBash = {...this.state.bugBash};
        bugBash.description = newDescription;
        this.updateState({bugBash: bugBash} as IBugBashEditorState);
    }

    private _updateDescriptionField(fieldRefName: string) {
        let bugBash = {...this.state.bugBash};
        bugBash.itemDescriptionField = fieldRefName;
        this.updateState({bugBash: bugBash} as IBugBashEditorState);
    }

    private _updateStartTime(newStartTime: Date) {
        let bugBash = {...this.state.bugBash};
        bugBash.startTime = newStartTime;
        this.updateState({bugBash: bugBash} as IBugBashEditorState);
    }

    private _updateEndTime(newEndTime: Date) {        
        let bugBash = {...this.state.bugBash};
        bugBash.endTime = newEndTime;
        this.updateState({bugBash: bugBash} as IBugBashEditorState);
    }

    private _updateAutoAccept(autoAccept: boolean) {        
        let bugBash = {...this.state.bugBash};
        bugBash.autoAccept = autoAccept;
        this.updateState({bugBash: bugBash} as IBugBashEditorState);
    }

    private _updateAcceptTemplate(templateId: string) {
        let bugBash = {...this.state.bugBash};
        bugBash.acceptTemplate = {team: VSS.getWebContext().team.id, templateId: templateId};
        this.updateState({bugBash: bugBash} as IBugBashEditorState);
    }
}