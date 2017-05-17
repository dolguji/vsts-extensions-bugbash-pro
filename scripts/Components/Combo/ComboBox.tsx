/// <reference types="react" />

import * as React from "react";
import * as Controls from "VSS/Controls";
import * as Combos from "VSS/Controls/Combos";
import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";

export interface IComboBoxProps extends IBaseComponentProps {
    value?: string;
    options?: Combos.IComboOptions;
    onChange: (newValue: string) => void;
}

export class ComboBox extends BaseComponent<IComboBoxProps, IBaseComponentState> {
    private _control: Combos.Combo;

    /**
     * Reference to the combo control's DOM
     */
    public refs: {
        [key: string]: (Element);
        container: (HTMLElement);
    };

    protected getDefaultClassName(): string {
        return "combobox bowtie";
    }

    public render(): JSX.Element {
        return <div ref="container" className={this.getClassName()}></div>;
    }

    public componentDidMount(): void {
        this._control = Controls.Control.create(Combos.Combo, $(this.refs.container), {...this.props.options || {}, change: () => {
            this.props.onChange(this._control.getText());
        }});

        this._control.setInputText(this.props.value || "");
    }

    public componentWillUnmount(): void {
        this._dispose();
    }

    public componentWillReceiveProps(nextProps: IComboBoxProps) {        
        this._control.setInputText(nextProps.value || "");
    }

    private _dispose(): void {
        if (this._control) {
            this._control.dispose();
            this._control = null;
        }
    }
}
