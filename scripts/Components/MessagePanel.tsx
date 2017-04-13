import * as React from "react";

export interface IMessagePanelProps {
    message: string;
    messageType: MessageType;
}

export enum MessageType {
    Error,
    Warning,
    Info,
    Success
}

export var MessagePanel: React.StatelessComponent<IMessagePanelProps> =
    (props: IMessagePanelProps): JSX.Element => {
        let className = "bowtie-fabric message-panel";
        let iconClass: string = "";

        switch (props.messageType) {
            case MessageType.Error:
                iconClass = "bowtie-icon bowtie-status-failure";
                className += " message-error";
                break;
            case MessageType.Warning:
                iconClass = "bowtie-icon bowtie-status-warning";
                className += " message-warning";
                break;
            case MessageType.Success:
                iconClass = "bowtie-icon bowtie-status-success-outline";
                className += " message-success";
                break;
            default:
                iconClass = "bowtie-icon bowtie-status-info";
                className += " message-info";
                break;
        }

        return (
            <div className={className}>
                <span className={iconClass} />    
                <span className="message-text">{props.message}</span>
            </div>
        );
}
