import "../../css/BugBashEditor.scss";

import * as React from "react";

import { TextField } from "OfficeFabric/TextField";
import { DatePicker } from "OfficeFabric/DatePicker";
import { Label } from "OfficeFabric/Label";
import { Dropdown, IDropdownOption, IDropdownProps } from "OfficeFabric/Dropdown";
import { autobind } from "OfficeFabric/Utilities";
import { Checkbox } from "OfficeFabric/Checkbox";

import { WorkItemTemplateReference, WorkItemField, WorkItemType, FieldType } from "TFS/WorkItemTracking/Contracts";
import Utils_String = require("VSS/Utils/String");
import Utils_Date = require("VSS/Utils/Date");
import Utils_Core = require("VSS/Utils/Core");

import { MessagePanel, MessageType } from "VSTS_Extension/Components/Common/MessagePanel";
import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { InputError } from "VSTS_Extension/Components/Common/InputError";
import { Loading } from "VSTS_Extension/Components/Common/Loading";
import { InfoLabel } from "VSTS_Extension/Components/Common/InfoLabel";
import { BaseStore } from "VSTS_Extension/Flux/Stores/BaseStore";
import { WorkItemFieldActions } from "VSTS_Extension/Flux/Actions/WorkItemFieldActions";
import { WorkItemTypeActions } from "VSTS_Extension/Flux/Actions/WorkItemTypeActions";
import { WorkItemTemplateActions } from "VSTS_Extension/Flux/Actions/WorkItemTemplateActions";

import { StoresHub } from "../Stores/StoresHub";
import { IBugBash } from "../Interfaces";
import { RichEditorComponent } from "./RichEditorComponent";

export interface IBugBashEditorProps extends IBaseComponentProps {
    error?: string;
    bugBash: IBugBash;
    onChange: (updatedBugBash: IBugBash) => void;
    save: () => void;
}

export interface IBugBashEditorState extends IBaseComponentState {
    loading?: boolean;
}

export class BugBashEditor extends BaseComponent<IBugBashEditorProps, IBugBashEditorState>  {
    private _imagePastedHandler: (event, data) => void;    

    constructor(props: IBugBashEditorProps, context?: any) {
        super(props, context);
        this._imagePastedHandler = Utils_Core.delegate(this, this._onImagePaste);
    }

    protected initializeState(): void {
        this.state = {
            loading: true
        };
    }

    protected getStores(): BaseStore<any, any, any>[] {
        return [StoresHub.workItemFieldStore, StoresHub.workItemTemplateStore, StoresHub.workItemTypeStore];
    }

    protected getStoresState(): IBugBashEditorState {
        return {
            loading: StoresHub.workItemTypeStore.isLoading() || StoresHub.workItemFieldStore.isLoading() || StoresHub.workItemTemplateStore.isLoading()
        } as IBugBashEditorState;
    }

    public componentDidMount(): void {
        super.componentDidMount();
        $(window).off("imagepasted", this._imagePastedHandler);
        $(window).on("imagepasted", this._imagePastedHandler);    
    
        WorkItemFieldActions.initializeWorkItemFields();
        WorkItemTypeActions.initializeWorkItemTypes();
        WorkItemTemplateActions.initializeWorkItemTemplates();
    }  

    public componentWillUnmount() {
        super.componentWillUnmount();
        $(window).off("imagepasted", this._imagePastedHandler);
    }      

    public render(): JSX.Element {
        if (this.state.loading) {
            return <Loading />;
        }

        return (
            <div className="bugbash-editor">                
                { this.props.error && <MessagePanel messageType={MessageType.Error} message={this.props.error} /> }
                { this._renderEditor() }                
            </div>
        );
    }

    private _renderEditor(): JSX.Element {
        const bugBash = this.props.bugBash;

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
        
        return <div className="bugbash-editor-contents" onKeyDown={this._onEditorKeyDown} tabIndex={0}>
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
                        onChange={(_ev: React.FormEvent<HTMLElement>, isChecked: boolean) => this._updateAutoAccept(isChecked) } />

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

    private _onImagePaste(_event, args) {
        args.callback(null);
    }

    @autobind
    private _onEditorKeyDown(e: React.KeyboardEvent<any>) {
        if (e.ctrlKey && e.keyCode === 83) {
            e.preventDefault();
            this.props.save();
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
        let filteredTemplates = StoresHub.workItemTemplateStore
            .getAll()
            .filter((t: WorkItemTemplateReference) => Utils_String.equals(t.workItemTypeName, this.props.bugBash.workItemType || "", true));

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

    private _onChange(updatedBugBash: IBugBash) {
        this.props.onChange(updatedBugBash);
    }

    private _updateTitle(newTitle: string) {
        let bugBash = {...this.props.bugBash};
        bugBash.title = newTitle;
        this._onChange(bugBash);
    }

    private _updateWorkItemType(newType: string) {
        let bugBash = {...this.props.bugBash};
        bugBash.workItemType = newType;
        bugBash.acceptTemplate = {team: VSS.getWebContext().team.id, templateId: ""};  // reset template
        this._onChange(bugBash);
    }

    private _updateDescription(newDescription: string) {
        let bugBash = {...this.props.bugBash};
        bugBash.description = newDescription;
        this._onChange(bugBash);
    }

    private _updateDescriptionField(fieldRefName: string) {
        let bugBash = {...this.props.bugBash};
        bugBash.itemDescriptionField = fieldRefName;
        this._onChange(bugBash);
    }

    private _updateStartTime(newStartTime: Date) {
        let bugBash = {...this.props.bugBash};
        bugBash.startTime = newStartTime;
        this._onChange(bugBash);
    }

    private _updateEndTime(newEndTime: Date) {        
        let bugBash = {...this.props.bugBash};
        bugBash.endTime = newEndTime;
        this._onChange(bugBash);
    }

    private _updateAutoAccept(autoAccept: boolean) {        
        let bugBash = {...this.props.bugBash};
        bugBash.autoAccept = autoAccept;
        this._onChange(bugBash);
    }

    private _updateAcceptTemplate(templateId: string) {
        let bugBash = {...this.props.bugBash};
        bugBash.acceptTemplate = {team: VSS.getWebContext().team.id, templateId: templateId};
        this._onChange(bugBash);
    }
}