import * as React from "react";
import { autobind } from "OfficeFabric/Utilities";
import { Label } from "OfficeFabric/Label";
import { TextField } from "OfficeFabric/TextField";
import { Button, ButtonType } from "OfficeFabric/Button";

import { WorkItemTemplate, WorkItem, FieldType } from "TFS/WorkItemTracking/Contracts";
import { WorkItemFormNavigationService } from "TFS/WorkItemTracking/Services";

import { IBaseProps, IBugBash } from "../Models";
import { MessagePanel, MessageType } from "./MessagePanel";
import Helpers = require("../Helpers");

export interface INewWorkItemCreatorProps extends IBaseProps {
    item: IBugBash;
    addWorkItem: (workItem: WorkItem) => void;
}

export interface INewWorkItemCreatorState {
    newWorkItemFieldValues: IDictionaryStringTo<string>;
    workItemError?: string;
    creating?: boolean;
}

export class NewWorkItemCreator extends React.Component<INewWorkItemCreatorProps, INewWorkItemCreatorState> {
    constructor(props: INewWorkItemCreatorProps, context: any) {
        super(props, context);
        
        this.state ={
            newWorkItemFieldValues: {},
            workItemError: "",
            creating: false
        };
    }

    public render(): JSX.Element {
        return (
            <div className="add-workitem-contents">
                { this.state.workItemError && <MessagePanel message={this.state.workItemError} messageType={MessageType.Error} />}
                <div className="header">
                    <Label className="add-workitem-label">Create work item</Label>                    
                </div>
                {this._getManualFieldsNode()}
                <Button className="create-new-button" disabled={!this._canCreateWorkItem()} buttonType={ButtonType.primary} onClick={this._onAddClick}>
                    {this.state.creating ? "Creating..." : "Create" }
                </Button>
            </div>
        );
    }

    @autobind 
    private async _onAddViaDialogClick() {
        let workItemNavSvc = await WorkItemFormNavigationService.getService();
        let initialFieldValues: IDictionaryStringTo<string> = {};        

        if (this.props.item.templateId) {     
            let templateFound = await this.props.context.actionsCreator.ensureTemplateItem(this.props.item.templateId);

            if (templateFound) {
                let template = this.props.context.stores.workItemTemplateItemStore.getItem(this.props.item.templateId);
                initialFieldValues = { ...template.fields };
            }            
        } 

        if (initialFieldValues["System.Tags-Add"]) {
            initialFieldValues["System.Tags"] = initialFieldValues["System.Tags-Add"];
        }
        delete initialFieldValues["System.Tags-Add"];
        delete initialFieldValues["System.Tags-Remove"];

        workItemNavSvc.openNewWorkItem(this.props.item.workItemType, initialFieldValues);
    }

    @autobind
    private async _onAddClick() {
        if (!this._canCreateWorkItem()) {
            return;
        }        
        
        this.setState({...this.state, creating: true, workItemError: null});

        let template: WorkItemTemplate = null;
        let newWorkItemFieldValues = {...this.state.newWorkItemFieldValues} || {};
        let templateFieldValues: IDictionaryStringTo<string> = {};        

        if (this.props.item.templateId) {     
            let templateFound = await this.props.context.actionsCreator.ensureTemplateItem(this.props.item.templateId);

            if (templateFound) {
                template = this.props.context.stores.workItemTemplateItemStore.getItem(this.props.item.templateId);
                templateFieldValues = { ...template.fields };
            }            
        } 

        let extraTags: string[] = Helpers.parseTags(templateFieldValues["System.Tags-Add"]);
        delete templateFieldValues["System.Tags-Add"];
        delete templateFieldValues["System.Tags"];
        delete templateFieldValues["System.Tags-Remove"];

        newWorkItemFieldValues = { ...templateFieldValues, ...newWorkItemFieldValues };
        const existingTags = Helpers.parseTags(newWorkItemFieldValues["System.Tags"]);
        newWorkItemFieldValues["System.Tags"] = existingTags.concat(Helpers.getBugBashTag(this.props.item.id)).concat(extraTags).join(";");
        
        try {
            let workItem = await Helpers.createWorkItem(this.props.item.workItemType, newWorkItemFieldValues);
            this.setState({...this.state, creating: false, workItemError: null, newWorkItemFieldValues: {}});
            this.props.addWorkItem(workItem);
        }
        catch (e) {
            this.setState({...this.state, creating: false, workItemError: e.message});
        }
    }

     @autobind
    private _canCreateWorkItem(): boolean {
        if (this.state.creating) {
            return false;
        }
        let newWorkItemFieldValues = this.state.newWorkItemFieldValues || {};
        for (let manualField of this.props.item.manualFields) {
            if (!newWorkItemFieldValues[manualField] || newWorkItemFieldValues[manualField].trim() === "") {
                return false;
            }
        }
        return true;
    }

    private _getManualFieldsNode(): React.ReactNode {
        let newWorkItemFieldValues = this.state.newWorkItemFieldValues || {};

        return this.props.item.manualFields.map((fieldRefName: string) => {
            const field = this.props.context.stores.workItemFieldStore.getItem(fieldRefName);
            if (field) {
                return (
                    <div className="manual-field-row" key={fieldRefName}>
                        { field.type !== FieldType.PlainText && field.type !== FieldType.Html && (
                        <TextField label={field.name} 
                            value={newWorkItemFieldValues[fieldRefName] || ""}
                            required={true} 
                            onChanged={(newValue: string) => this._setWorkItemFieldValue(fieldRefName, newValue)} />
                        )}
                        { field.type === FieldType.PlainText || field.type === FieldType.Html && (
                        <div>
                            <Label required={true}>{field.name}</Label>
                            <div>
                                <div ref={(container: HTMLElement) => this._renderRichEditor(container, fieldRefName)}/>
                            </div>
                        </div>
                        )}
                    </div>
                );
            }
            else {
                return <MessagePanel message={`The field ${fieldRefName} doesn't exist anymore in your account.`} messageType={MessageType.Warning} />                
            }        
        });
    }

    @autobind
    private _renderRichEditor(container: HTMLElement, fieldRefName: string) {
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
                    this._setWorkItemFieldValue(fieldRefName, $(container).summernote('code'));
                }
            }
        });

        $(container).summernote('code', this.state.newWorkItemFieldValues[fieldRefName] || "");
    }

    @autobind
    private _setWorkItemFieldValue(fieldRefName: string, value: string) {
        let newWorkItemFieldValues = {...this.state.newWorkItemFieldValues};
        newWorkItemFieldValues[fieldRefName] = value;
        this.setState({...this.state, newWorkItemFieldValues: newWorkItemFieldValues});
    }    
}