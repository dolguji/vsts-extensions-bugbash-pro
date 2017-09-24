import * as React from "react";
import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "MB/Components/BaseComponent";
import { UIActionsHub } from "MB/Flux/Actions/ActionsHub";
import { Label } from "OfficeFabric/Label";
import { autobind } from "OfficeFabric/Utilities";

interface IBugBashResultCountState extends IBaseComponentState {
    count: number;
}

interface IBugBashResultCountProps extends IBaseComponentProps {
    count: number;
}

export class BugBashResultCount extends BaseComponent<IBugBashResultCountProps, IBugBashResultCountState> {
    protected initializeState() {
        this.state = {
            count: this.props.count
        };
    }

    public componentDidMount() {
        super.componentDidMount();
        UIActionsHub.OnGridItemCountChanged.addListener(this._resetCount);
    }  

    public componentWillUnmount() {
        super.componentWillUnmount();
        UIActionsHub.OnGridItemCountChanged.removeListener(this._resetCount);
    }

    public render(): JSX.Element {
       return <Label style={{ display: "flex", padding: 10, alignItems: "center"}}>{this.state.count} results</Label>;
    }

    @autobind
    private _resetCount(count: number) {
        this.updateState({count: count});
    }
}
