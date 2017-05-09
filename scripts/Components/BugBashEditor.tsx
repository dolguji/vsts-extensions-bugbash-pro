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

import { StoresHub } from "../Stores/StoresHub";
import { IBugBash, UrlActions } from "../Models";
import { BugBash } from "../Models/BugBash";

export interface IBugBashEditorProps extends IBaseComponentProps {
    id?: string;
}

export interface IBugBashEditorState extends IBaseComponentState {
    loading?: boolean;
    model?: IBugBash;
    error?: string;
}

export class BugBashEditor extends BaseComponent<IBugBashEditorProps, IBugBashEditorState>  {
    private _item: BugBash;
    private _richEditorContainer: JQuery;

    protected initializeState(): void {
        if (this.props.id) {
            this._item = new BugBash(StoresHub.bugBashStore.getItem(this.props.id));
        }
        else {
            this._item = BugBash.getNew();
        }

        this.state = {
            model: this._item.getModel(),
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
        this._item.attachChanged(() => {
            this.onStoreChanged();
        }); 
    }

    protected onStoreChanged() {
        this.updateState({
            model: this._item.getModel(),
            loading: StoresHub.workItemTypeStore.isLoaded() && StoresHub.workItemFieldStore.isLoaded() && StoresHub.workItemTemplateStore.isLoaded() ? false : true
        });
    }    

    public render(): JSX.Element {
        if (this.state.loading) {
            return <Loading />;
        }

        let model = this.state.model;
        let menuitems: IContextualMenuItem[] = [
            {
                key: "save", name: "Save", title: "Save", iconProps: {iconName: "Save"}, disabled: !this._item.isDirty() || !this._item.isValid(),
                onClick: async (event?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem) => {
                    if (this._item.isNew() && this._item.isDirty() && this._item.isValid()) {
                        let savedBugBash = await StoresHub.bugBashStore.createItem(model);
                        let navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
                        navigationService.updateHistoryEntry(UrlActions.ACTION_EDIT, { id: savedBugBash.id }, true);
                    }
                    else if (!this._item.isNew() && this._item.isDirty() && this._item.isValid()) {
                        let result = await StoresHub.bugBashStore.updateItem(model);                        

                        if (!result) {
                            this.updateState({error: "The bug bash version does not match with the latest version. Please refresh the page and try again."});
                        }
                        else {
                            this._item.renew(result);
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
                    this._richEditorContainer.summernote('code', this._item.getModel().description);
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
                        
                        await StoresHub.bugBashItemStore.clearAllItems(model.id);
                        await StoresHub.bugBashStore.deleteItem(model);
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
                <div className="editor-view-menu">
                    <CommandBar 
                        items={menuitems} 
                        farItems={[{
                            key: "Home", name: "Home", title: "Return to home view", iconProps: {iconName: "Home"}, 
                            onClick: async (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                                let navigationService: HostNavigationService = await VSS.getService(VSS.ServiceIds.Navigation) as HostNavigationService;
                                navigationService.updateHistoryEntry(UrlActions.ACTION_ALL, null);
                            }
                        }]} />
                </div>
                { this.state.error && (<MessageBar messageBarType={MessageBarType.error}>{this.state.error}</MessageBar> )}
                <div className="editor-view-contents">                    
                    <div className="first-section">                        
                        <TextField 
                            label='Title' 
                            value={model.title} 
                            onChanged={(newValue: string) => this._item.updateTitle(newValue)} 
                            onGetErrorMessage={this._getTitleError} />

                        <Label>Description</Label>

                        <div>
                            <div ref={this._renderRichEditor}/>
                        </div>
                    </div>
                    <div className="second-section">
                        <div className="checkbox-container">                            
                            <Checkbox 
                                className="auto-accept"
                                label=""
                                checked={model.autoAccept}
                                onChange={(ev: React.FormEvent<HTMLElement>, isChecked: boolean) => this._item.updateAutoAccept(isChecked) } />

                            <InfoLabel label="Auto Accept?" info="Auto create work items on creation of a bug bash item" />
                        </div>

                        <DatePicker 
                            label="Start Date" 
                            allowTextInput={true} 
                            isRequired={false} 
                            value={model.startTime}
                            onSelectDate={(newValue: Date) => this._item.updateStartTime(newValue)} />

                        <DatePicker 
                            label="Finish Date" 
                            allowTextInput={true}
                            isRequired={false} 
                            value={model.endTime} 
                            onSelectDate={(newValue: Date) => this._item.updateEndTime(newValue)} />

                        { model.startTime && model.endTime && Utils_Date.defaultComparer(model.startTime, model.endTime) >= 0 &&  (<InputError error="Bugbash end time cannot be a date before bugbash start time." />)}

                        <InfoLabel label="Work item type" info="Select a work item type which would be used to create work items for each bug bash item" />
                        <Dropdown 
                            className={!model.workItemType ? "editor-dropdown no-margin" : "editor-dropdown"}
                            onRenderList={this._onRenderCallout} 
                            required={true} 
                            options={witItems}                            
                            onChanged={(option: IDropdownOption) => this._item.updateWorkItemType(option.key as string)} />

                        { !model.workItemType && (<InputError error="A work item type is required." />) }

                        <InfoLabel label="Description field" info="Select a HTML field that you would want to set while creating a workitem for each bug bash item" />
                        <Dropdown 
                            className={!model.itemDescriptionField ? "editor-dropdown no-margin" : "editor-dropdown"}
                            onRenderList={this._onRenderCallout} 
                            required={true} 
                            options={fieldItems} 
                            onChanged={(option: IDropdownOption) => this._item.updateDescriptionField(option.key as string)} />

                        { !model.itemDescriptionField && (<InputError error="A description field is required." />) }       

                        <InfoLabel label="Work item template" info="Select a work item template that would be applied during work item creation." />
                        <Dropdown 
                            className="editor-dropdown"
                            onRenderList={this._onRenderCallout} 
                            options={this._getTemplateDropdownOptions(model.acceptTemplate.templateId)} 
                            onChanged={(option: IDropdownOption) => this._item.updateAcceptTemplate(option.key as string)} />
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
    private _renderRichEditor(container: HTMLElement) {
        this._richEditorContainer = $(container);
        this._richEditorContainer.summernote({
            height: 260,
            minHeight: 260,
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
                onChange: () => {
                    this._item.updateDescription(this._richEditorContainer.summernote('code'));
                }
            }
        });

        this._richEditorContainer.summernote('code', this.state.model.description);
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
}