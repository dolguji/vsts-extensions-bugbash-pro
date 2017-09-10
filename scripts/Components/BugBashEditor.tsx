import "../../css/BugBashEditor.scss";

import * as React from "react";

import { TextField } from "OfficeFabric/TextField";
import { DatePicker } from "OfficeFabric/DatePicker";
import { Label } from "OfficeFabric/Label";
import { autobind } from "OfficeFabric/Utilities";
import { Checkbox } from "OfficeFabric/Checkbox";
import { MessageBar, MessageBarType } from "OfficeFabric/MessageBar";
import { Dropdown, IDropdownOption, IDropdownProps } from "OfficeFabric/Dropdown";

import { WebApiTeam } from "TFS/Core/Contracts";
import { WorkItemTemplateReference, WorkItemField, WorkItemType, FieldType } from "TFS/WorkItemTracking/Contracts";

import { ComboBox } from "MB/Components/Combobox";
import { DateUtils } from "MB/Utils/Date";
import { CoreUtils } from "MB/Utils/Core";
import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "MB/Components/BaseComponent";
import { InputError } from "MB/Components/InputError";
import { Loading } from "MB/Components/Loading";
import { InfoLabel } from "MB/Components/InfoLabel";
import { BaseStore } from "MB/Flux/Stores/BaseStore";
import { WorkItemFieldActions } from "MB/Flux/Actions/WorkItemFieldActions";
import { WorkItemTypeActions } from "MB/Flux/Actions/WorkItemTypeActions";
import { WorkItemTemplateActions } from "MB/Flux/Actions/WorkItemTemplateActions";
import { TeamActions } from "MB/Flux/Actions/TeamActions";

import { StoresHub } from "../Stores/StoresHub";
import { RichEditorComponent } from "./RichEditorComponent";
import { BugBash } from "../ViewModels/BugBash";
import { ErrorKeys, BugBashFieldNames, SizeLimits } from "../Constants";
import { BugBashErrorMessageActions } from "../Actions/BugBashErrorMessageActions";
import { copyImageToGitRepo } from "../Helpers";

export interface IBugBashEditorProps extends IBaseComponentProps {
    bugBash: BugBash;
}

export interface IBugBashEditorState extends IBaseComponentState {
    templates: IDropdownOption[];
    error?: string;
    teamNames?: string[];
    workItemTypeNames?: string[];
    htmlFieldNames?: string[];
}

export class BugBashEditor extends BaseComponent<IBugBashEditorProps, IBugBashEditorState>  {
    private _imagePastedHandler: (event, data) => void;    
    private _updateBugBashDelayedFunction: CoreUtils.DelayedFunction;

    constructor(props: IBugBashEditorProps, context?: any) {
        super(props, context);
        this._imagePastedHandler = CoreUtils.delegate(this, this._onImagePaste);
    }

    protected initializeState() {
        this.state = {
            loading: true,
            templates: [],
            teamNames: null,
            workItemTypeNames: null,
            htmlFieldNames: null
        };
    }

    protected getStores(): BaseStore<any, any, any>[] {
        return [StoresHub.workItemFieldStore, StoresHub.teamStore, StoresHub.workItemTemplateStore, StoresHub.workItemTypeStore, StoresHub.bugBashErrorMessageStore];
    }

    protected getStoresState(): IBugBashEditorState {
        let state = {
            loading: StoresHub.workItemTypeStore.isLoading() || StoresHub.workItemFieldStore.isLoading() || StoresHub.teamStore.isLoading(),
            error: StoresHub.bugBashErrorMessageStore.getItem(ErrorKeys.BugBashError)
        } as IBugBashEditorState;  

        if (StoresHub.workItemTypeStore.isLoaded() && !this.state.workItemTypeNames) {
            state = {...state, workItemTypeNames: StoresHub.workItemTypeStore.getAll().map(w => w.name)};
        }

        if (StoresHub.workItemFieldStore.isLoaded() && !this.state.htmlFieldNames) {    
            state = {...state, htmlFieldNames: StoresHub.workItemFieldStore.getAll().filter(f => f.type === FieldType.Html).map(f => f.name)};        
        }

        if (StoresHub.teamStore.isLoaded() && !this.state.teamNames) {       
            state = {...state, teamNames: StoresHub.teamStore.getAll().map(t => t.name)};
        }              

        const acceptTemplateTeam = this.props.bugBash.getFieldValue<string>(BugBashFieldNames.AcceptTemplateTeam);
        if (StoresHub.workItemTemplateStore.isLoaded(acceptTemplateTeam)) {
            const templates = StoresHub.workItemTemplateStore.getItem(acceptTemplateTeam).map((template: WorkItemTemplateReference) => {
                return {
                    key: template.id.toLowerCase(),
                    text: template.name
                }
            });

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
        TeamActions.initializeTeams();

        const acceptTemplateTeam = this.props.bugBash.getFieldValue<string>(BugBashFieldNames.AcceptTemplateTeam);
        if (acceptTemplateTeam) {
            WorkItemTemplateActions.initializeWorkItemTemplates(acceptTemplateTeam);
        }
    }  

    public componentWillReceiveProps(nextProps: Readonly<IBugBashEditorProps>) {
        const nextAcceptTemplateTeam = nextProps.bugBash.getFieldValue<string>(BugBashFieldNames.AcceptTemplateTeam);

        if (nextAcceptTemplateTeam) {
            if (StoresHub.workItemTemplateStore.isLoaded(nextAcceptTemplateTeam)) {
                const templates = StoresHub.workItemTemplateStore.getItem(nextAcceptTemplateTeam).map((template: WorkItemTemplateReference) => {
                    return {
                        key: template.id.toLowerCase(),
                        text: template.name
                    }
                });
    
                this.updateState({templates: templates} as IBugBashEditorState);
            }
            else {
                WorkItemTemplateActions.initializeWorkItemTemplates(nextAcceptTemplateTeam);
            }            
        }
        else {
            this.updateState({templates: []} as IBugBashEditorState);
        }
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
    }

    private _renderEditor(): JSX.Element {
        const bugBash = this.props.bugBash;        
        const bugBashTitle = bugBash.getFieldValue<string>(BugBashFieldNames.Title);
        const workItemTypeName = bugBash.getFieldValue<string>(BugBashFieldNames.WorkItemType);
        const bugBashDescription = bugBash.getFieldValue<string>(BugBashFieldNames.Description);
        const startTime = bugBash.getFieldValue<Date>(BugBashFieldNames.StartTime);
        const endTime = bugBash.getFieldValue<Date>(BugBashFieldNames.EndTime);
        const itemDescriptionField = bugBash.getFieldValue<string>(BugBashFieldNames.ItemDescriptionField);
        const acceptTemplateId = bugBash.getFieldValue<string>(BugBashFieldNames.AcceptTemplateId);
        const acceptTemplateTeamId = bugBash.getFieldValue<string>(BugBashFieldNames.AcceptTemplateTeam);

        const team = StoresHub.teamStore.getItem(acceptTemplateTeamId);
        const workItemType = StoresHub.workItemTypeStore.getItem(workItemTypeName);
        const field = StoresHub.workItemFieldStore.getItem(itemDescriptionField);

        return <div className="bugbash-editor-contents" onKeyDown={this._onEditorKeyDown} tabIndex={0}>
            <div className="first-section">                        
                <TextField 
                    label="Title"
                    value={bugBashTitle} 
                    onChanged={(newValue: string) => this._onChange(BugBashFieldNames.Title, newValue)} />
                
                { (bugBashTitle == null || bugBashTitle.trim() === "") && <InputError error="Title is required." /> }
                { bugBashTitle && bugBashTitle.length > SizeLimits.TitleFieldMaxLength && <InputError error={`The length of the title should be less than 257 characters, actual is ${bugBashTitle.length}.`} /> }

                <Label>Description</Label>
                <RichEditorComponent 
                    containerId="bugbash-description-editor" 
                    data={bugBashDescription} 
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
                    onChange={(newValue: string) => this._onChange(BugBashFieldNames.Description, newValue)} />
            </div>
            <div className="second-section">                
                <DatePicker 
                    label="Start Date" 
                    allowTextInput={true} 
                    isRequired={false} 
                    value={startTime}
                    onSelectDate={(newValue: Date) => this._onChange(BugBashFieldNames.StartTime, newValue)} />

                <DatePicker 
                    label="Finish Date" 
                    allowTextInput={true}
                    isRequired={false} 
                    value={endTime} 
                    onSelectDate={(newValue: Date) => this._onChange(BugBashFieldNames.EndTime, newValue)} />

                { startTime && endTime && DateUtils.defaultComparer(startTime, endTime) >= 0 &&  <InputError error="Bugbash end time cannot be a date before bugbash start time." />}

                <InfoLabel label="Work item type" info="Select a work item type which would be used to create work items for each bug bash item" />
                <ComboBox                             
                    value={workItemType ? workItemType.name : workItemTypeName} 
                    options={{
                        type: "list",
                        mode: "drop",
                        allowEdit: true,
                        source: this.state.workItemTypeNames
                    }} 
                    error={this._getWorkItemTypeError(workItemTypeName, workItemType)}
                    onChange={this._onWorkItemTypeChange}/>

                <InfoLabel label="Description field" info="Select a HTML field that you would want to set while creating a workitem for each bug bash item" />
                <ComboBox                             
                    value={field ? field.name : itemDescriptionField} 
                    options={{
                        type: "list",
                        mode: "drop",
                        allowEdit: true,
                        source: this.state.htmlFieldNames
                    }} 
                    error={this._getFieldError(itemDescriptionField, field)}
                    onChange={this._onFieldChange}/>
            </div>
            <div className="third-section">
                <div className="checkbox-container">                            
                    <Checkbox 
                        className="auto-accept"
                        label=""
                        checked={bugBash.getFieldValue<boolean>(BugBashFieldNames.AutoAccept)}
                        onChange={(_ev: React.FormEvent<HTMLElement>, isChecked: boolean) => this._onChange(BugBashFieldNames.AutoAccept, isChecked, true) } />

                    <InfoLabel label="Auto Accept?" info="Auto create work items on creation of a bug bash item" />
                </div>

                <InfoLabel label="Template Team" info="Select a team to pull its templates." />
                <ComboBox                             
                    value={team ? team.name : acceptTemplateTeamId} 
                    options={{
                        type: "list",
                        mode: "drop",
                        allowEdit: true,
                        source: this.state.teamNames
                    }} 
                    error={this._getTeamError(acceptTemplateTeamId, team)}
                    onChange={this._onTeamChange}/>

                <InfoLabel label="Work item template" info="Select a work item template that would be applied during work item creation." />
                <Dropdown 
                    selectedKey={acceptTemplateId.toLowerCase()}
                    onRenderList={this._onRenderCallout}
                    options={this.state.templates} 
                    onChanged={(option: IDropdownOption) => this._onChange(BugBashFieldNames.AcceptTemplateId, option.key as string, true)} />

                { !acceptTemplateId && <InputError error="A work item template is required." /> }
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
        copyImageToGitRepo(args.data, "Description", args.callback);        
    }    
    
    @autobind
    private _onEditorKeyDown(e: React.KeyboardEvent<any>) {
        if (e.ctrlKey && e.keyCode === 83) {
            e.preventDefault();
            this.props.bugBash.save();
        }
    }

    private _onChange<T>(fieldName: BugBashFieldNames, fieldValue: T, immediate?: boolean) {
        if (this._updateBugBashDelayedFunction) {
            this._updateBugBashDelayedFunction.cancel();
        }

        if (immediate) {
            this.props.bugBash.setFieldValue<T>(fieldName, fieldValue);
        }
        else {
            this._updateBugBashDelayedFunction = CoreUtils.delay(this, 200, () => {
                this.props.bugBash.setFieldValue<T>(fieldName, fieldValue);
            });
        }
    }

    @autobind
    private _onTeamChange(newTeamName: string) {
        this.props.bugBash.setFieldValue<string>(BugBashFieldNames.AcceptTemplateId, "", false);

        const newTeam = StoresHub.teamStore.getItem(newTeamName);
        this._onChange(BugBashFieldNames.AcceptTemplateTeam, newTeam ? newTeam.id : newTeamName);
    }

    private _getTeamError(teamId: string, team: WebApiTeam): string {
        if (teamId == null || teamId.trim() === "") {
            return "A template team is required.";
        }
        else if (team == null) {
            return `Team "${teamId}" does not exist.`;
        }

        return null;
    }

    @autobind
    private _onFieldChange(newFieldName: string) {
        const newField = StoresHub.workItemFieldStore.getItem(newFieldName);
        this._onChange(BugBashFieldNames.ItemDescriptionField, newField ? newField.referenceName : newFieldName);
    }

    private _getFieldError(fieldRefName: string, field: WorkItemField): string {
        if (fieldRefName == null || fieldRefName.trim() === "") {
            return "A description field is required.";
        }
        else if (field == null || field.type !== FieldType.Html) {
            return `Field "${fieldRefName}" does not exist or is not a html field.`;
        }

        return null;
    }

    @autobind
    private _onWorkItemTypeChange(newTypeName: string) {
        const newWIT = StoresHub.workItemTypeStore.getItem(newTypeName);
        this._onChange(BugBashFieldNames.WorkItemType, newWIT ? newWIT.name : newTypeName);
    }

    private _getWorkItemTypeError(workItemTypeName: string, workItemType: WorkItemType): string {
        if (workItemTypeName == null || workItemTypeName.trim() === "") {
            return "A work item type is required.";
        }
        else if (workItemType == null) {
            return `Work item type "${workItemTypeName}" does not exist.`;
        }

        return null;
    }
}