import "../../css/BugBashItemView.scss";

import * as React from "react";
import { autobind } from "OfficeFabric/Utilities";
import { Label } from "OfficeFabric/Label";
import { TextField } from "OfficeFabric/TextField";
import { PrimaryButton } from "OfficeFabric/Button";
import { MessageBar, MessageBarType } from "OfficeFabric/MessageBar";

import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { BaseStore } from "VSTS_Extension/Stores/BaseStore";
import { Loading } from "VSTS_Extension/Components/Common/Loading";

import { WorkItemFormNavigationService } from "TFS/WorkItemTracking/Services";

import { IBugBashItemDocument, ICommentDocument } from "../Models";
import Helpers = require("../Helpers");
import { BugBashItemCommentStore } from "../Stores/BugBashItemCommentStore";
import { StoresHub } from "../Stores/StoresHub";
import { BugBashItem } from "../Models/BugBashItem";
import { RichEditor } from "./RichEditor/RichEditor";

export interface IBugBashItemViewProps extends IBaseComponentProps {
    id?: string;
    bugBashId: string;
}

export interface IBugBashItemViewState extends IBaseComponentState {
    model?: IBugBashItemDocument;
    comments?: ICommentDocument[];
    loadingComments?: boolean;
    workItemError?: string;
    saving?: boolean;
}

export class BugBashItemView extends BaseComponent<IBugBashItemViewProps, IBugBashItemViewState> {
    private _item: BugBashItem;

    protected getStoresToLoad(): {new (): BaseStore<any, any, any>}[] {
        return [BugBashItemCommentStore];
    }

    protected initializeState() {
        if (this.props.id) {
            this._item = new BugBashItem(StoresHub.bugBashItemStore.getItem(this.props.id));
        }
        else {
            this._item = BugBashItem.getNew(this.props.bugBashId);
        }

        const model = this._item.getModel();

        this.state = {
            model: model,
            comments: [],
            loadingComments: true
        };
    }

    protected initialize() {
        if (this._item) {
            this._item.detachChanged();        
        }

        this._initialize();
    }   

    protected onStoreChanged() {
        if (this._item) {
            this.updateState({
                comments: StoresHub.bugBashItemCommentStore.getItems(this._item.getModel().id),
                loadingComments: !StoresHub.bugBashItemCommentStore.isDataLoaded(this._item.getModel().id)
            });
        }        
    }

    public componendWillUnmount() {
        super.componentWillUnmount();
        if (this._item) {
            this._item.detachChanged();        
        }
    }

    public componentWillReceiveProps(nextProps: Readonly<IBugBashItemViewProps>): void {
        if (this._item) {
            this._item.detachChanged();        
        }

        if (nextProps.id) {
            this._item = new BugBashItem(StoresHub.bugBashItemStore.getItem(nextProps.id));
        }
        else {
            this._item = BugBashItem.getNew(nextProps.bugBashId);
        }

        this.updateState({
            model: this._item.getModel(),
            comments: [],
            loadingComments: true
        });

        this._initialize();
    }

    private _initialize() {        
        const model = this._item.getModel();
        if (model.id) {
            StoresHub.bugBashItemCommentStore.refreshItems(model.id);
        }
        else {
            this.updateState({loadingComments: false});
        }

        this._item.attachChanged(() => {
            this.updateState({model: this._item.getModel()});
        }); 
    }

    public render(): JSX.Element {
        let model = this.state.model;
        if (!model) {
            return <Loading />;
        }

        return (
            <div className="add-workitem-contents">
                { this.state.workItemError && <MessageBar messageBarType={MessageBarType.error}>{this.state.workItemError}</MessageBar>}

                <div className="header">
                    <Label className="add-workitem-label">Create work item</Label>
                </div>

                <TextField label="Title" 
                        value={model.title}
                        required={true} 
                        onChanged={(newValue: string) => this._item.updateTitle(newValue)} />

                <div>
                    <Label className="item-description">Description</Label>
                    <RichEditor containerId="rich-editor" data={model.description} onChange={(newValue: string) => this._item.updateDescription(newValue)} />
                </div>

                <PrimaryButton className="create-new-button" disabled={this.state.saving || !this._item.isDirty() || !this._item.isValid()}  onClick={this._onAddClick}>
                    {this.state.saving ? "Saving..." : "Save" }
                </PrimaryButton>
            </div>
        );
    }

    @autobind
    private async _onAddClick() {
        if (this.state.model.id) {
            StoresHub.bugBashItemStore.updateItem(this.state.model);
        }
        else {
            StoresHub.bugBashItemStore.createItem(this.state.model);
        }
    }
}