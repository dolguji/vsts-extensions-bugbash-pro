import "../../css/BugBashDetails.scss";

import * as React from "react";

import { autobind } from "OfficeFabric/Utilities";
import { MessageBar, MessageBarType } from "OfficeFabric/MessageBar";

import { CoreUtils } from "MB/Utils/Core";
import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "MB/Components/BaseComponent";
import { Loading } from "MB/Components/Loading";
import { BaseStore } from "MB/Flux/Stores/BaseStore";

import { LongTextActions } from "../Actions/LongTextActions";
import { StoresHub } from "../Stores/StoresHub";
import { ErrorKeys } from "../Constants";
import { BugBashErrorMessageActions } from "../Actions/BugBashErrorMessageActions";
import { copyImageToGitRepo } from "../Helpers";
import { RichEditorComponent } from "./RichEditorComponent";
import { LongText } from "../ViewModels/LongText";

export interface IBugBashDetailsProps extends IBaseComponentProps {
    id: string;
    isEditMode: boolean;
}

export interface IBugBashDetailsState extends IBaseComponentState {
    error: string;
    longText: LongText;
}

export class BugBashDetails extends BaseComponent<IBugBashDetailsProps, IBugBashDetailsState>  {
    private _imagePastedHandler: (event, data) => void;    
    private _updateDelayedFunction: CoreUtils.DelayedFunction;

    constructor(props: IBugBashDetailsProps, context?: any) {
        super(props, context);
        this._imagePastedHandler = CoreUtils.delegate(this, this._onImagePaste);
    }

    protected initializeState() {
        this.state = {
            loading: true,
            longText: null,
            error: null
        };
    }

    protected getStores(): BaseStore<any, any, any>[] {
        return [StoresHub.longTextStore, StoresHub.bugBashErrorMessageStore];
    }

    protected getStoresState(): IBugBashDetailsState {
        return {
            loading: StoresHub.longTextStore.isLoading(this.props.id),
            longText: StoresHub.longTextStore.getItem(this.props.id),
            error: StoresHub.bugBashErrorMessageStore.getItem(ErrorKeys.BugBashDetailsError)
        } as IBugBashDetailsState;  
    }

    public componentDidMount() {
        super.componentDidMount();
        $(window).off("imagepasted", this._imagePastedHandler);
        $(window).on("imagepasted", this._imagePastedHandler);    
    
        LongTextActions.initializeLongText(this.props.id);
    }

    public componentWillUnmount() {
        super.componentWillUnmount();        

        $(window).off("imagepasted", this._imagePastedHandler);
        this._dismissErrorMessage();
    }

    public render(): JSX.Element {
        if (this.state.loading) {
            return <Loading />;
        }

        return <div className="bugbash-details">
            { this.state.error && 
                <MessageBar 
                    className="message-panel"
                    messageBarType={MessageBarType.error} 
                    onDismiss={this._dismissErrorMessage}>
                    {this.state.error}
                </MessageBar>
            }
            
            <div
                className="bugbash-details-contents" 
                onKeyDown={this._onEditorKeyDown} 
                tabIndex={0}>

                { this.props.isEditMode && 
                    <RichEditorComponent 
                        containerId="bugbash-description-editor" 
                        data={this.state.longText.Text} 
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
                        onChange={this._onChange} />
                }

                { !this.props.isEditMode && 
                    <div className="bugbash-details-html" dangerouslySetInnerHTML={{__html: this.state.longText.Text}}></div>
                }
            </div>
        </div>;
    }

    @autobind
    private _dismissErrorMessage() {
        BugBashErrorMessageActions.dismissErrorMessage(ErrorKeys.BugBashDetailsError);
    }

    @autobind
    private _onChange(newValue: string) {
        if (this._updateDelayedFunction) {
            this._updateDelayedFunction.cancel();
        }

        let longText = this.state.longText;
        this._updateDelayedFunction = CoreUtils.delay(this, 200, () => {
            longText.setDetails(newValue);
        });
    }

    private async _onImagePaste(_event, args) {
        try {
            const imageUrl = await copyImageToGitRepo(args.data, "Details");   
            args.callback(imageUrl);
        }
        catch (e) {
            BugBashErrorMessageActions.showErrorMessage(e, ErrorKeys.BugBashDetailsError);
        }
    }    
    
    @autobind
    private _onEditorKeyDown(e: React.KeyboardEvent<any>) {
        if (e.ctrlKey && e.keyCode === 83) {
            e.preventDefault();
            this.state.longText.save();
        }
    }
}