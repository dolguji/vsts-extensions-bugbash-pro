import "../../css/BugBashEditor.scss";

import * as React from "react";

import { TextField } from "OfficeFabric/TextField";
import { DatePicker } from "OfficeFabric/DatePicker";
import { Label } from "OfficeFabric/Label";
import { autobind } from "OfficeFabric/Utilities";
import { Checkbox } from "OfficeFabric/Checkbox";
import { MessageBar, MessageBarType } from "OfficeFabric/MessageBar";
import { Dropdown, IDropdownOption, IDropdownProps } from "OfficeFabric/Dropdown";

import { WorkItemTemplateReference, WorkItemField, WorkItemType, FieldType } from "TFS/WorkItemTracking/Contracts";
import Utils_String = require("VSS/Utils/String");
import Utils_Date = require("VSS/Utils/Date");
import { delegate, delay, DelayedFunction } from "VSS/Utils/Core";

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
import { BugBash } from "../ViewModels/BugBash";
import { ErrorKeys } from "../Constants";
import { BugBashErrorMessageActions } from "../Actions/BugBashErrorMessageActions";

export interface IBugBashEditorProps extends IBaseComponentProps {
    bugBash: BugBash;
}

export interface IBugBashEditorState extends IBaseComponentState {
    workItemTypes: IDropdownOption[];
    htmlFields: IDropdownOption[];
    templates: IDropdownOption[];
    error?: string;
}

export class BugBashEditor extends BaseComponent<IBugBashEditorProps, IBugBashEditorState>  {
    private _imagePastedHandler: (event, data) => void;    
    private _updateBugBashDelayedFunction: DelayedFunction;

    constructor(props: IBugBashEditorProps, context?: any) {
        super(props, context);
        this._imagePastedHandler = delegate(this, this._onImagePaste);
    }

    protected initializeState() {
        this.state = {
            loading: true,
            workItemTypes: null,
            htmlFields: null,
            templates: null
        };
    }

    protected getStores(): BaseStore<any, any, any>[] {
        return [StoresHub.workItemFieldStore, StoresHub.workItemTemplateStore, StoresHub.workItemTypeStore, StoresHub.bugBashErrorMessageStore];
    }

    protected getStoresState(): IBugBashEditorState {
        let state = {
            loading: StoresHub.workItemTypeStore.isLoading() || StoresHub.workItemFieldStore.isLoading() || StoresHub.workItemTemplateStore.isLoading(),
            error: StoresHub.bugBashErrorMessageStore.getItem(ErrorKeys.BugBashError)
        } as IBugBashEditorState;

        if (StoresHub.workItemTypeStore.isLoaded() && !this.state.workItemTypes) {
            const workItemTypes: IDropdownOption[] = StoresHub.workItemTypeStore.getAll().map((workItemType: WorkItemType) => {
                return {
                    key: workItemType.name.toLowerCase(),
                    text: workItemType.name
                }
            });
            state = {...state, workItemTypes: workItemTypes};
        }

        if (StoresHub.workItemFieldStore.isLoaded() && !this.state.htmlFields) {            
            const htmlFields: IDropdownOption[] = StoresHub.workItemFieldStore.getAll().filter(f => f.type === FieldType.Html).map((field: WorkItemField) => {
                return {
                    key: field.referenceName.toLowerCase(),
                    text: field.name
                }
            });
            state = {...state, htmlFields: htmlFields};
        }

        if (StoresHub.workItemTemplateStore.isLoaded() && !this.state.templates) {            
            let emptyTemplateItem = [
                {   
                    key: "", text: "<No template>"
                }
            ];
            let filteredTemplates = StoresHub.workItemTemplateStore
                .getAll()
                .filter((t: WorkItemTemplateReference) => Utils_String.equals(t.workItemTypeName, this.props.bugBash.updatedModel.workItemType || "", true));

            const templates = emptyTemplateItem.concat(filteredTemplates.map((template: WorkItemTemplateReference) => {
                return {
                    key: template.id.toLowerCase(),
                    text: template.name
                }
            }));
            state = {...state, templates: templates};
        }

        return state;
    }

    public componentDidMount() {
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

    public componentWillReceiveProps(nextProps: IBugBashEditorProps) {
        if (nextProps.bugBash.updatedModel.workItemType != this.props.bugBash.updatedModel.workItemType) {
            let emptyTemplateItem = [
                {   
                    key: "", text: "<No template>"
                }
            ];
            let filteredTemplates = StoresHub.workItemTemplateStore
                .getAll()
                .filter((t: WorkItemTemplateReference) => Utils_String.equals(t.workItemTypeName, this.props.bugBash.updatedModel.workItemType || "", true));

            const templates = emptyTemplateItem.concat(filteredTemplates.map((template: WorkItemTemplateReference) => {
                return {
                    key: template.id.toLowerCase(),
                    text: template.name
                }
            }));

            this.updateState({templates: templates} as IBugBashEditorState);
        }
    }

    public render(): JSX.Element {
        if (this.state.loading) {
            return <Loading />;
        }

        return (
            <div className="bugbash-editor">
                { this.state.error && 
                    <MessageBar 
                        className="message-panel"
                        messageBarType={MessageBarType.error} 
                        onDismiss={this._dismissErrorMessage}>
                        {this.state.error}
                    </MessageBar>
                }
                { this._renderEditor() }                
            </div>
        );
    }

    @autobind
    private _dismissErrorMessage() {
        BugBashErrorMessageActions.dismissErrorMessage(ErrorKeys.BugBashError);
    };

    private _renderEditor(): JSX.Element {
        const bugBash = this.props.bugBash;        

        return <div className="bugbash-editor-contents" onKeyDown={this._onEditorKeyDown} tabIndex={0}>
            <div className="first-section">                        
                <TextField 
                    label='Title' 
                    value={bugBash.updatedModel.title} 
                    onChanged={this._updateTitle} />
                
                { (bugBash.updatedModel.title == null || bugBash.updatedModel.title.trim() === "") && <InputError error="Title is required." /> }
                { bugBash.updatedModel.title && bugBash.updatedModel.title.length > 256 && <InputError error={`The length of the title should be less than 257 characters, actual is ${bugBash.updatedModel.title.length}.`} /> }

                <Label>Description</Label>
                <RichEditorComponent 
                    containerId="bugbash-description-editor" 
                    data={bugBash.updatedModel.description} 
                    editorOptions={{
                        svgPath: `${VSS.getExtensionContext().baseUri}/css/libs/icons.png`,
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
                    onChange={this._updateDescription} />
            </div>
            <div className="second-section">
                <div className="checkbox-container">                            
                    <Checkbox 
                        className="auto-accept"
                        label=""
                        checked={bugBash.updatedModel.autoAccept}
                        onChange={(_ev: React.FormEvent<HTMLElement>, isChecked: boolean) => this._updateAutoAccept(isChecked) } />

                    <InfoLabel label="Auto Accept?" info="Auto create work items on creation of a bug bash item" />
                </div>

                <DatePicker 
                    label="Start Date" 
                    allowTextInput={true} 
                    isRequired={false} 
                    value={bugBash.updatedModel.startTime}
                    onSelectDate={this._updateStartTime} />

                <DatePicker 
                    label="Finish Date" 
                    allowTextInput={true}
                    isRequired={false} 
                    value={bugBash.updatedModel.endTime} 
                    onSelectDate={this._updateEndTime} />

                { bugBash.updatedModel.startTime && bugBash.updatedModel.endTime && Utils_Date.defaultComparer(bugBash.updatedModel.startTime, bugBash.updatedModel.endTime) >= 0 &&  <InputError error="Bugbash end time cannot be a date before bugbash start time." />}

                <InfoLabel label="Work item type" info="Select a work item type which would be used to create work items for each bug bash item" />
                <Dropdown 
                    selectedKey={bugBash.updatedModel.workItemType.toLowerCase()}
                    onRenderList={this._onRenderCallout}
                    className={!bugBash.updatedModel.workItemType ? "editor-dropdown no-margin" : "editor-dropdown"}
                    required={true} 
                    options={this.state.workItemTypes}                            
                    onChanged={(option: IDropdownOption) => this._updateWorkItemType(option.key as string)} />

                { !bugBash.updatedModel.workItemType && <InputError error="A work item type is required." /> }

                <InfoLabel label="Description field" info="Select a HTML field that you would want to set while creating a workitem for each bug bash item" />
                <Dropdown 
                    selectedKey={bugBash.updatedModel.itemDescriptionField.toLowerCase()}
                    onRenderList={this._onRenderCallout}
                    className={!bugBash.updatedModel.itemDescriptionField ? "editor-dropdown no-margin" : "editor-dropdown"}
                    required={true} 
                    options={this.state.htmlFields} 
                    onChanged={(option: IDropdownOption) => this._updateDescriptionField(option.key as string)} />

                { !bugBash.updatedModel.itemDescriptionField && <InputError error="A description field is required." /> }     

                <InfoLabel label="Work item template" info="Select a work item template that would be applied during work item creation." />
                <Dropdown 
                    selectedKey={bugBash.updatedModel.acceptTemplate.templateId.toLowerCase()}
                    onRenderList={this._onRenderCallout}
                    className="editor-dropdown"
                    options={this.state.templates} 
                    onChanged={(option: IDropdownOption) => this._updateAcceptTemplate(option.key as string)} />
            </div>
        </div>;
    }

    @autobind
    private _onRenderCallout(props?: IDropdownProps, defaultRender?: (props?: IDropdownProps) => JSX.Element): JSX.Element {
        return <div className="callout-container">
                {defaultRender(props)}
            </div>;        
    }

    private _onImagePaste(_event, args) {
        args.callback(null);
    }

    @autobind
    private _onEditorKeyDown(e: React.KeyboardEvent<any>) {
        if (e.ctrlKey && e.keyCode === 83) {
            e.preventDefault();
            this.props.bugBash.save();
        }
    }

    private _onChange(updatedModel: IBugBash) {
        if (this._updateBugBashDelayedFunction) {
            this._updateBugBashDelayedFunction.cancel();
        }

        this._updateBugBashDelayedFunction = delay(this, 200, () => {
            this.props.bugBash.update(updatedModel);
        });
    }

    @autobind
    private _updateTitle(newTitle: string) {
        let updatedModel = {...this.props.bugBash.updatedModel};
        updatedModel.title = newTitle;
        this._onChange(updatedModel);
    }

    private _updateWorkItemType(newType: string) {
        let updatedModel = {...this.props.bugBash.updatedModel};
        updatedModel.workItemType = newType;
        updatedModel.acceptTemplate = {team: VSS.getWebContext().team.id, templateId: ""};  // reset template
        this._onChange(updatedModel);
    }

    @autobind
    private _updateDescription(newDescription: string) {
        let updatedModel = {...this.props.bugBash.updatedModel};
        updatedModel.description = newDescription;
        this._onChange(updatedModel);
    }

    private _updateDescriptionField(fieldRefName: string) {
        let updatedModel = {...this.props.bugBash.updatedModel};
        updatedModel.itemDescriptionField = fieldRefName;
        this._onChange(updatedModel);
    }

    @autobind
    private _updateStartTime(newStartTime: Date) {
        let updatedModel = {...this.props.bugBash.updatedModel};
        updatedModel.startTime = newStartTime;
        this._onChange(updatedModel);
    }

    @autobind
    private _updateEndTime(newEndTime: Date) {        
        let updatedModel = {...this.props.bugBash.updatedModel};
        updatedModel.endTime = newEndTime;
        this._onChange(updatedModel);
    }

    private _updateAutoAccept(autoAccept: boolean) {        
        let updatedModel = {...this.props.bugBash.updatedModel};
        updatedModel.autoAccept = autoAccept;
        this._onChange(updatedModel);
    }

    private _updateAcceptTemplate(templateId: string) {
        let updatedModel = {...this.props.bugBash.updatedModel};
        updatedModel.acceptTemplate = {team: VSS.getWebContext().team.id, templateId: templateId};
        this._onChange(updatedModel);
    }
}