import * as React from "react";
import { TextField } from "OfficeFabric/TextField";
import { CommandBar } from "OfficeFabric/CommandBar";
import { DatePicker } from "OfficeFabric/DatePicker";
import { Label } from "OfficeFabric/Label";
import { Dropdown } from "OfficeFabric/components/Dropdown/Dropdown";
import { IDropdownOption, IDropdownProps } from "OfficeFabric/components/Dropdown/Dropdown.Props";
import { IContextualMenuItem } from "OfficeFabric/components/ContextualMenu/ContextualMenu.Props";
import { TagPicker, ITag } from 'OfficeFabric/components/pickers/TagPicker/TagPicker';
import { autobind } from "OfficeFabric/Utilities";

import { HostNavigationService } from "VSS/SDK/Services/Navigation";
import { WorkItemTemplateReference, WorkItemField, WorkItemType } from "TFS/WorkItemTracking/Contracts";
import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");
import Utils_Date = require("VSS/Utils/Date");

import { IBugBash, IBaseProps, LoadingState, UrlActions, Constants } from "../Models";
import { BugBash } from "../BugBash";
import { Loading } from "./Loading";
import { isInteger } from "../Helpers";
import { InputError } from "./InputError";

import { MessagePanel, MessageType } from "./MessagePanel";

export interface IBugBashEditorProps extends IBaseProps {
    id?: string;
}

export interface IBugBashEditorState {
    loadingState: LoadingState;
    model: IBugBash;
    templates: WorkItemTemplateReference[];
    fields: WorkItemField[];    
    workItemTypes: WorkItemType[];
    error?: string;
}

export class BugBashEditor extends React.Component<IBugBashEditorProps, IBugBashEditorState> {
    private _item: BugBash;
    
    constructor(props: IBugBashEditorProps, context: any) {
        super(props, context);

        if (props.id) {
            this._item = new BugBash(this.props.context.stores.bugBashItemStore.getItem(props.id));
        }
        else {
            this._item = BugBash.getNew();
        } 

        this.state = {
            model: this._item.getModel(),
            fields: [],
            workItemTypes: [],
            templates: [],
            loadingState: LoadingState.Loading,
            error: null
        };
    }

    public componentDidMount() {   
        this.props.context.stores.workItemTypeStore.addChangedListener(this._onStoreChanged);
        this.props.context.stores.workItemFieldStore.addChangedListener(this._onStoreChanged);
        this.props.context.stores.workItemTemplateStore.addChangedListener(this._onStoreChanged);
        this._item.attachChanged(() => {
            this._onStoreChanged();
        });  

        this._initialize();      
    }

    public componentWillUnmount() {
        this.props.context.stores.workItemTypeStore.removeChangedListener(this._onStoreChanged);
        this.props.context.stores.workItemFieldStore.removeChangedListener(this._onStoreChanged);
        this.props.context.stores.workItemTemplateStore.removeChangedListener(this._onStoreChanged);    
        this._item.detachChanged();
    }

    private _initialize(): void {
        this.props.context.actionsCreator.initializeWorkItemFields();
        this.props.context.actionsCreator.initializeWorkItemTemplates();
        this.props.context.actionsCreator.initializeWorkItemTypes();
    }

    private _onStoreChanged = (handler?: IEventHandler): void => {
        let newState = this._getStateFromStore();
        this.setState(newState);
    }    

    private _getStateFromStore(): IBugBashEditorState {
        return {
            fields: this.props.context.stores.workItemFieldStore.getAll(),
            templates: this.props.context.stores.workItemTemplateStore.getAll(),
            workItemTypes: this.props.context.stores.workItemTypeStore.getAll(),
            model: this._item.getModel(),
            loadingState: this.props.context.stores.workItemTypeStore.isLoaded() && this.props.context.stores.workItemFieldStore.isLoaded() && this.props.context.stores.workItemTemplateStore.isLoaded() ? LoadingState.Loaded : LoadingState.Loading
        };
    }

    public render(): JSX.Element {
        if (!this.state || this.state.loadingState === LoadingState.Loading) {
            return <Loading />;
        }

        let model = this.state.model;
        let menuitems: IContextualMenuItem[] = [
            {
                key: "save", name: "Save", title: "Save", iconProps: {iconName: "Save"}, disabled: !this._item.isDirty() || !this._item.isValid(),
                onClick: async (event?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem) => {
                    if (this._item.isNew() && this._item.isDirty() && this._item.isValid()) {
                        let savedBugBash = await this.props.context.actionsCreator.createBugBash(model);
                        let navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
                        navigationService.updateHistoryEntry(UrlActions.ACTION_EDIT, { id: savedBugBash.id }, true);
                    }
                    else if (!this._item.isNew() && this._item.isDirty() && this._item.isValid()) {
                        let result = await this.props.context.actionsCreator.updateBugBash(model);
                        this._item.renew();

                        if (!result) {
                            this.setState({...this.state, error: "The bug bash version does not match with the latest version. Please refresh the page and try again."})
                        }
                    }
                }
            },
            {
                key: "undo", name: "Undo", title: "Undo changes", iconProps: {iconName: "Undo"}, disabled: this._item.isNew() || !this._item.isDirty(),
                onClick: async (event?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem) => {
                    let dialogService: IHostDialogService = await VSS.getService(VSS.ServiceIds.Dialog) as IHostDialogService;
                    try {
                        await dialogService.openMessageDialog("Are you sure you want to undo your changes to this instance?", { useBowtieStyle: true });            
                    }
                    catch (e) {
                        // user selected "No"" in dialog
                        return;
                    }

                    this._item.reset();
                }
            },
            {
                key: "delete", name: "Delete", title: "Delete", iconProps: {iconName: "Delete"}, disabled: this._item.isNew(),
                onClick: async (event?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem) => {
                    if (!this._item.isNew()) {
                        let dialogService: IHostDialogService = await VSS.getService(VSS.ServiceIds.Dialog) as IHostDialogService;
                        try {
                            await dialogService.openMessageDialog("Are you sure you want to delete this instance?", { useBowtieStyle: true });            
                        }
                        catch (e) {
                            // user selected "No"" in dialog
                            return;
                        }
                        
                        await this.props.context.actionsCreator.deleteBugBash(model);
                        let navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
                        navigationService.updateHistoryEntry(UrlActions.ACTION_ALL, null);
                    }
                }
            },
            {
                key: "results", name: "Show results", title: "Show results", iconProps: {iconName: "ShowResults"}, disabled: this._item.isNew(),
                onClick: async (event?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem) => {
                    if (!this._item.isNew()) {
                        if (this._item.isDirty()) {
                            let dialogService: IHostDialogService = await VSS.getService(VSS.ServiceIds.Dialog) as IHostDialogService;
                            try {
                                await dialogService.openMessageDialog("Are you sure you want to go back to results view? This action will reset your unsaved changes.", { useBowtieStyle: true });            
                            }
                            catch (e) {
                                // user selected "No"" in dialog
                                return;
                            }
                        }
                        this._item.reset();
                        let navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
                        navigationService.updateHistoryEntry(UrlActions.ACTION_VIEW, {id: model.id});
                    }
                }
            },
        ];        

        let witItems: IDropdownOption[] = this.state.workItemTypes.map((workItemType: WorkItemType, index: number) => {
            return {
                key: workItemType.name,
                index: index + 1,
                text: workItemType.name,
                selected: model.workItemType ? Utils_String.equals(model.workItemType, workItemType.name, true) : false
            }
        });

        let tagPickerClassName = "field-picker";
        if (model.manualFields.length == 0) {
            tagPickerClassName += " invalid";
        }

        return (
            <div className="editor-view">
                <div className="editor-view-menu">
                    <CommandBar items={menuitems} 
                        farItems={[{
                            key: "Home", name: "Home", title: "Return to home view", iconProps: {iconName: "Home"}, 
                            onClick: async (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                                let navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
                                navigationService.updateHistoryEntry(UrlActions.ACTION_ALL, null);
                            }
                        }]} />
                </div>
                <div className="editor-view-contents">
                    <div className="first-section">
                        { this.state.error && (<MessagePanel message={this.state.error} messageType={MessageType.Error} /> )}
                        <TextField label='Title' required={true} value={model.title} onChanged={(newValue: string) => this._item.updateTitle(newValue)} onGetErrorMessage={this._getTitleError} />
                        <Label>Description</Label>
                        <div>
                            <div ref={this._renderRichEditor}/>
                        </div>
                    </div>
                    <div className="section-wrapper">
                        <div className="second-section">                    
                            <DatePicker label="Start Date" allowTextInput={true} isRequired={false} value={model.startTime} onSelectDate={(newValue: Date) => this._item.updateStartTime(newValue)} />
                            <DatePicker label="Finish Date" allowTextInput={true} isRequired={false} value={model.endTime} onSelectDate={(newValue: Date) => this._item.updateEndTime(newValue)} />
                            { model.startTime && model.endTime && Utils_Date.defaultComparer(model.startTime, model.endTime) >= 0 &&  (<InputError error="Bugbash end time cannot be a date before bugbash start time." />)}
                        </div>
                        <div className="third-section">
                            <Dropdown label="Work item type" disabled={!this._item.isNew()} onRenderList={this._onRenderCallout} required={true} options={witItems} onChanged={(option: IDropdownOption) => this._item.updateWorkItemType(option.key as string)} />                            
                            { !model.workItemType && (<InputError error="A work item type is required." />) }

                            <Dropdown label="Work item template" onRenderList={this._onRenderCallout} options={this._getTemplateDropdownOptions(model.templateId)} onChanged={(option: IDropdownOption) => this._item.updateTemplate(option.key as string)} />
                            <Label required={true}>Manually entered fields</Label>
                            <TagPicker className={tagPickerClassName}
                                defaultSelectedItems={model.manualFields.map(f => this._getFieldTag(f))}
                                onResolveSuggestions={this._onFieldFilterChanged}
                                getTextFromItem={item => item.name}
                                onChange={items => this._item.updateManualFields(items.map(item => item.key))}
                                pickerSuggestionsProps={
                                    {
                                        suggestionsHeaderText: 'Suggested Fields',
                                        noResultsFoundText: 'No fields Found'
                                    }
                                }
                            />
                            { model.manualFields.length == 0 && (<InputError error="Atleast one field must be manually entered." />) }
                        </div>
                        <div className="fourth-section">
                            <Dropdown label="Accept Work item template" onRenderList={this._onRenderCallout} 
                                options={this._getTemplateDropdownOptions(model.configTemplates[Constants.ACCEPT_CONFIG_TEMPLATE_KEY])} 
                                onChanged={(option: IDropdownOption) => this._item.updateConfigTemplate(Constants.ACCEPT_CONFIG_TEMPLATE_KEY, option.key as string)} />

                            <Dropdown label="Reject Work item template" onRenderList={this._onRenderCallout} 
                                options={this._getTemplateDropdownOptions(model.configTemplates[Constants.REJECT_CONFIG_TEMPLATE_KEY])} 
                                onChanged={(option: IDropdownOption) => this._item.updateConfigTemplate(Constants.REJECT_CONFIG_TEMPLATE_KEY, option.key as string)} />
                        </div>
                    </div>
                </div>
            </div>
        );
    }

    @autobind
    private _onRenderCallout(props?: IDropdownProps, defaultRender?: (props?: IDropdownProps) => JSX.Element): JSX.Element {
        return (
            <div className="callout-container">
                {defaultRender(props)}
            </div>
        );
    }

    @autobind
    private _getTemplateDropdownOptions(selectedValue: string): IDropdownOption[] {
        let emptyTemplateItem = [
            {   
                key: "", index: 0, text: "<No template>", 
                selected: !selectedValue
            }
        ];
        let filteredTemplates = this.state.templates.filter((t: WorkItemTemplateReference) => Utils_String.equals(t.workItemTypeName, this.state.model.workItemType));
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
    private _renderRichEditor(container: HTMLElement) {
        $(container).summernote({
            height: 200,
            minHeight: 200,
            toolbar: [
                // [groupName, [list of button]]
                ['style', ['bold', 'italic', 'underline', 'clear']],
                ['fontsize', ['fontsize']],
                ['color', ['color']],
                ['para', ['ul', 'ol', 'paragraph']],
                ['insert', ['link', 'picture']],
                ['fullscreen', ['fullscreen']]
            ],
            callbacks: {
                onBlur: () => {
                    this._item.updateDescription($(container).summernote('code'));
                }
            }
        });

        $(container).summernote('code', this.state.model.description);
    }

    @autobind
    private _getFieldTag(refName: string): ITag {
        return {
            key: refName,
            name: Utils_Array.first(this.state.fields, field => Utils_String.equals(field.referenceName, refName, true)).name
        }
    }

    @autobind
    private _onFieldFilterChanged(filterText: string, tagList: ITag[]): ITag[] {
        return filterText
            ? this.state.fields.filter(field => field.name.toLowerCase().indexOf(filterText.toLowerCase()) === 0 
                && Utils_Array.findIndex(tagList, (tag: ITag) => Utils_String.equals(tag.key, field.referenceName, true)) === -1).map(field => {
                    return { key: field.referenceName, name: field.name};
                }) 
            : [];
    }

    @autobind
    private _getTitleError(value: string): string | IPromise<string> {
        if (!value) {
            return "Title is required";
        }
        if (value.length > 128) {
            return `The length of the title should less than 128 characters, actual is ${value.length}.`
        }
        return "";
    }

    @autobind
    private _getTagError(value: string): string | IPromise<string> {
        if (!value) {
            return "Tag is required";
        }
        if (value.length > 128) {
            return `The length of the tag should less than 128 characters, actual is ${value.length}.`
        }
        return "";
    }
}