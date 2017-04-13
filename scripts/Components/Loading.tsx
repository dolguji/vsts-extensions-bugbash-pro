import * as React from "react";
import { Spinner, SpinnerType } from "OfficeFabric/Spinner";

export var Loading: React.StatelessComponent<any> = (): JSX.Element => {
    return (
        <div className="content-loading">
            <Spinner className="loading-spinner" type={SpinnerType.large} label="Loading..." />
        </div>
    );
};
