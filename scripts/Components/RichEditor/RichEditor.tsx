import * as React from "react";

import "trumbowyg/dist/trumbowyg";
import "./PasteImagePlugin";
import "trumbowyg/dist/ui/trumbowyg.min.css";

export interface IRichEditorProps {
    containerId: string;
    placeholder?: string
    data: string;
    onChange: (newValue: string) => void;
    editorOptions?: any;
}

export class RichEditor extends React.Component<IRichEditorProps, {}> {
    private _richEditorContainer: JQuery;

    constructor(props: IRichEditorProps, context) {
        super();
        ($ as any).trumbowyg.svgPath = "/css/libs/icons.svg";
    }

    public componentDidMount() {
        this._richEditorContainer = $("#" + this.props.containerId);
        this._richEditorContainer.trumbowyg(this.props.editorOptions || {})
        .on("tbwchange", () => this.props.onChange(this._richEditorContainer.trumbowyg("html")));

        this._richEditorContainer.trumbowyg("html", this.props.data);
    }

    public componentWillUnmount() {
        this._richEditorContainer.trumbowyg('destroy');
    }

    public componentWillReceiveProps(nextProps: IRichEditorProps) {
        if (nextProps.data !== this._richEditorContainer.trumbowyg("html")) {
            this._richEditorContainer.trumbowyg('html', nextProps.data);
        }        
    }

    public render() {
        return (
            <div id={this.props.containerId} className="rich-editor" placeholder={this.props.placeholder || ""}>
                
            </div>
        );
    }
}