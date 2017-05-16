/// <reference types="react" />

import * as React from "react";
import * as Controls from "VSS/Controls";
import * as Combos from "VSS/Controls/Combos";
import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";

export interface IComboBoxProps extends IBaseComponentProps {
    options: Combos.IComboOptions;
}

export class ComboBox extends BaseComponent<IComboBoxProps, IBaseComponentState> {
    private _control: Combos.Combo;
    private _onInputFocusInProgress: boolean;

    /**
     * Reference to the web access combo control's DOM
     */
    public refs: {
        [key: string]: (Element);
        container: (HTMLElement);
    };

    protected getDefaultClassName(): string {
        return "combobox";
    }

    /**
     * Render the container with the given classname which wrap's the web access combo control
     */
    public render(): JSX.Element {
        return <div ref="container" className={this.getClassName()}></div>;
    }

    /**
     * Create the web access combo control during this react control mount
     */
    public componentDidMount(): void {
        this._control = this._createControl(this.refs.container, this.props);
    }

    /**
     * Remove the web access control during this react control unmount
     */
    public componentWillUnmount(): void {
        this._dispose();
    }

    /**
     * Set enabled state for this combo control when react control is updated
     */
    public componentDidUpdate(): void {
        this._control.setEnabled(this.props.options.enabled);
    }

    /**
     * Dispose and delete the reference of the web access control
     */
    private _dispose(): void {
        if (this._control) {
            // Web access controls have dispose
            this._control.dispose();
            this._control = null;
        }
    }

    /**
     * Creates the web access combo control
     */
    private _createControl(element: HTMLElement, props: IComboBoxProps): Combos.Combo {
        let control = Controls.Control.create(Combos.Combo, $(element), props.options);

        // Disabled style is associated with the child but border is set for the parent
        // and due to limitation of styling parent based on child class, we are adding our own class here
        if (props.options.enabled) {
            control._element.addClass("enabled-border");
        } else {
            control._element.addClass("disabled-border");
        }

        return control;
    }
}
