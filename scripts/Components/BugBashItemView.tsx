import "../../css/BugBashItemView.scss";

import * as React from "react";
import { autobind } from "OfficeFabric/Utilities";
import { Label } from "OfficeFabric/Label";
import { TextField } from "OfficeFabric/TextField";
import { PrimaryButton } from "OfficeFabric/Button";
import { MessageBar, MessageBarType } from "OfficeFabric/MessageBar";

import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { BaseStore } from "VSTS_Extension/Stores/BaseStore";

import { WorkItemFormNavigationService } from "TFS/WorkItemTracking/Services";

import { IBugBashItemDocument, ICommentDocument } from "../Models";
import Helpers = require("../Helpers");
import { BugBashItemCommentStore } from "../Stores/BugBashItemCommentStore";
import { StoresHub } from "../Stores/StoresHub";
import { BugBashItem } from "../Models/BugBashItem";

export interface IBugBashItemViewProps extends IBaseComponentProps {
    itemModel: IBugBashItemDocument;
    bugBashId: string;
}

export interface IBugBashItemViewState extends IBaseComponentState {
    item?: BugBashItem;
    comments?: ICommentDocument[];
    loadingComments?: boolean;
    workItemError?: string;
    saving?: boolean;
}

export class BugBashItemView extends BaseComponent<IBugBashItemViewProps, IBugBashItemViewState> {
    protected getStoresToLoad(): {new (): BaseStore<any, any, any>}[] {
        return [BugBashItemCommentStore];
    }

    protected initializeState() {
        this.state = {
            item: this.props.itemModel && this.props.itemModel.id ? new BugBashItem(this.props.itemModel) : BugBashItem.getNew(this.props.bugBashId),
            comments: [],
            loadingComments: true
        };
    }

    protected initialize() {
        if (this.props.itemModel && this.props.itemModel.id) {
            StoresHub.bugBashItemCommentStore.refreshItems(this.props.itemModel && this.props.itemModel.id);
        }
        else {
            this.updateState({loadingComments: false});
        }
    }   

    protected onStoreChanged() {
        if (this.props.itemModel && this.props.itemModel.id) {
            this.updateState({
                comments: StoresHub.bugBashItemCommentStore.getItems(this.props.itemModel.id),
                loadingComments: !StoresHub.bugBashItemCommentStore.isDataLoaded(this.props.itemModel.id)
            });
        }        
    }

    public componentWillReceiveProps(nextProps: Readonly<IBugBashItemViewProps>): void {
        this.updateState({
            item: nextProps.itemModel && nextProps.itemModel.id ? new BugBashItem(nextProps.itemModel) : BugBashItem.getNew(nextProps.bugBashId),
            comments: [],
            loadingComments: true
        });

        if (nextProps.itemModel && nextProps.itemModel.id) {
            StoresHub.bugBashItemCommentStore.refreshItems(this.props.itemModel && this.props.itemModel.id);
        }
        else {
            this.updateState({loadingComments: false});
        }
    }

    public render(): JSX.Element {
        let model = this.state.item.getModel();

        return (
            <div className="add-workitem-contents">
                { this.state.workItemError && <MessageBar messageBarType={MessageBarType.error}>{this.state.workItemError}</MessageBar>}

                <div className="header">
                    <Label className="add-workitem-label">Create work item</Label>
                </div>

                <TextField label="Title" 
                        value={model.title}
                        required={true} 
                        onChanged={(newValue: string) => this.state.item.updateTitle(newValue)} />

                <div>
                    <Label className="item-description">Description</Label>
                    <div ref={(container: HTMLElement) => this._renderRichEditor(container)}/>
                </div>

                <PrimaryButton className="create-new-button" disabled={!this.state.item.isDirty() || !this.state.item.isValid()}  onClick={this._onAddClick}>
                    {this.state.saving ? "Saving..." : "Save" }
                </PrimaryButton>
            </div>
        );
    }

    @autobind
    private async _onAddClick() {

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
                    this.state.item.updateDescription($(container).summernote('code'));
                }
            }
        });

        $(container).summernote('code', this.state.item.getModel().description || "");
    }
}