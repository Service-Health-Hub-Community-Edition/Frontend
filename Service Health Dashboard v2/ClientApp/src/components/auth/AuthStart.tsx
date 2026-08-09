import * as React from 'react';
import * as msTeams from '@microsoft/teams-js';
import { MsalContext, useMsal } from "@azure/msal-react";
import { loginRequest } from '../../auth/authConfig';
import { RedirectRequest } from '@azure/msal-browser';
import { Text } from '@fluentui/react';

const request : RedirectRequest = {
    scopes: loginRequest.scopes,
    redirectUri: '/auth-end'
}


export default class AuthStart extends React.Component {
    static contextType = MsalContext;

    componentDidMount(){
        if (this.context.accounts.length == 0) {
            this.context.instance.loginRedirect(request);
        }
    }

    public render() {
        return (
            <div style={{ height: "calc(100vh - 72px)" }}>
                <div className="authNotification" style={{
                    display: 'flex',
                    justifyContent: 'center',
                    alignItems: 'center',
                    width: '100%',
                    height: '100%'
                }}>
                    <div className="authNotificationContents" style={{ textAlign: 'center' }}>
                        <Text variant={'xxLarge'}>Signing you in</Text><br />
                        <Text variant='medium'>Just a few moments and we'll get you there</Text>
                    </div>
                </div>
            </div>
        );
    }
}
