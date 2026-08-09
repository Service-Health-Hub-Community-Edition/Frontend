import * as React from 'react';
import { Text } from '@fluentui/react';
import { Separator } from '@fluentui/react';

export interface IAccessDeniedState {
    message: string;
}

export class AccessDenied extends React.Component<{ message?: string }, IAccessDeniedState> {
    constructor(props: { message?: string }) {
        super(props);

        this.state = {
            message: ""
        };
    }

    public render() {
        const {
            message } = this.state;

        return (
            <div style = {{ height: "calc(100vh - 72px)"}}>
                <div className="accessDenied" style={{
                    display: 'flex',
                    justifyContent: 'center',
                    alignItems: 'center',
                    width: '100%',
                    height: '100%'
                }}>
                    <div className="accessDeniedContents" style={{ textAlign: 'center' }}>
                        <img src='/images/access-denied.svg' width='140px' /><br /><br />
                        <Text variant={'xxLarge'}>That didn't work.</Text><br />
                        <Text variant='medium'>We're sorry, but you don't have access to this resource.</Text>
                    </div>
                </div>
            </div>
        );
    }
}