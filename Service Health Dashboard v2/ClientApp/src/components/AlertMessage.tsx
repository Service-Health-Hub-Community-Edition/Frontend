import * as React from 'react';
import { Text, Dialog, DialogType, DialogFooter, PrimaryButton } from '@fluentui/react';

export interface IAlertState {
    title: string;
    header?: string;
    message?: string;
    technicalDetails?: string;
    isOpen: boolean;
    isBlocking: boolean;
}

export class AlertMessage extends React.Component<{ title: string, header?: string, message?: string, technicalDetails?: string, isOpen: boolean, isBlocking: boolean, onClose?: () => void }, IAlertState> {
    constructor(props: { title: string, header?: string, message?: string, technicalDetails?: string, isOpen: boolean, isBlocking: boolean, onClose?: () => void }) {
        super(props);

        this.state = {
            title: props.title,
            header: props.header,
            message: props.message,
            technicalDetails: props.technicalDetails,
            isOpen: props.isOpen,
            isBlocking: props.isBlocking
        };
    }

    public render() {
        const {
            title, header, message, technicalDetails, isOpen, isBlocking } = this.state;

        return (
            <Dialog
                hidden={!isOpen}
                onDismiss={this._closeDialog}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: title,
                }}
                modalProps={{
                    isBlocking: isBlocking
                }}
                styles={{
                    main: {
                        selectors: {
                            ['@media (min-width: 480px)']: {
                                width: 640,
                                minWidth: 300,
                                maxWidth: '1000px'
                            }
                        }
                    }
                }}
            >

                {header ? (
                    <Text variant='xLarge'>
                        <br/>{header}<br /><br/>
                    </Text>) : ""}

                {message ? (
                    <Text variant='medium'>
                        <div dangerouslySetInnerHTML={{ __html: message! }} /><br /><br />
                    </Text>) : (
                    <Text variant='medium'>
                        No further information provided.<br /><br />
                    </Text>
                )}

                {technicalDetails ? (
                    <div>
                        <Text variant='medium'>
                            <b>Technical details:</b><br />
                        </Text>
                        <Text variant='medium'>
                            <b>{technicalDetails}</b><br />
                        </Text>
                    </div>) : ""}

                {this.props.children}

                <DialogFooter>
                    <PrimaryButton onClick={this._closeDialog} text="Close" />
                </DialogFooter>
            </Dialog>

        );
    }

    componentDidMount() {
        this.setState({
            title: this.props.title,
            header: this.props.header,
            message: this.props.message,
            technicalDetails: this.props.technicalDetails,
            isOpen: this.props.isOpen,
            isBlocking: this.props.isBlocking
        });
    }

    componentDidUpdate(previousProps: any, previousState: IAlertState) {
        if (previousProps.title !== this.props.title ||
            previousProps.header !== this.props.header ||
            previousProps.message !== this.props.message ||
            previousProps.technicalDetails !== this.props.technicalDetails ||
            previousProps.isOpen !== this.props.isOpen ||
            previousProps.isBlocking !== this.props.isBlocking)
                this.setState({
                    title: this.props.title,
                    header: this.props.header,
                    message: this.props.message,
                    technicalDetails: this.props.technicalDetails,
                    isOpen: this.props.isOpen,
                    isBlocking: this.props.isBlocking
                });
    }

    private _closeDialog = (): void => {
        if (this.props.onClose)
            this.props.onClose();
    }
}