import * as React from 'react';
import * as msTeams from '@microsoft/teams-js';
import { AuthenticatedTemplate, MsalAuthenticationTemplate, MsalContext, UnauthenticatedTemplate } from "@azure/msal-react";
import { InteractionType, SsoSilentRequest } from '@azure/msal-browser';
import { loginRequest } from '../../auth/authConfig';
import { checkInTeams } from './detectTeams';
import { Text } from '@fluentui/react';

export interface IAuthenticationTemplateState {
    inTeams: boolean;
    ssoSilentRequest: SsoSilentRequest;
    authenticatedInTeams: boolean;
}

export default class AuthenticationTemplate extends React.Component < {}, IAuthenticationTemplateState> {
    static contextType = MsalContext;
    constructor(props: any) {
        super(props);
        this.state = {
            authenticatedInTeams: false,
            inTeams: checkInTeams(),
            ssoSilentRequest: {
                scopes: ["User.Read", "email", "openid", "profile", "offline_access"],
                loginHint: ''
            }
        };
        this.updateState = this.updateState.bind(this);
    }

    componentDidMount(){
        if (!this.state.ssoSilentRequest.loginHint){
            msTeams.getContext(context => {
                const ssoSilentRequest = {...this.state.ssoSilentRequest};
                ssoSilentRequest.loginHint = context.loginHint as string;
                this.setState({ssoSilentRequest});
            });
        }
    }

    async componentDidUpdate(prevProps: {}, prevState:IAuthenticationTemplateState){
        const self = this;
        if (this.state.ssoSilentRequest.loginHint && prevState.ssoSilentRequest.loginHint !== this.state.ssoSilentRequest.loginHint) {
            this.context.instance.ssoSilent(this.state.ssoSilentRequest).then((result: any) => {
                self.forceUpdate();
            }).catch((error: any) => {
                msTeams.initialize();
                msTeams.authentication.authenticate({
                    url: "/auth-start",
                    width: 600,
                    height: 535,
                    successCallback: function (result) {
                        self.updateState();
                    },
                    failureCallback: function (result) {alert(result);}
                });
            });
        }
        if (this.state.authenticatedInTeams) {
            await this.context.instance.ssoSilent(this.state.ssoSilentRequest);
        }

    }
    public render(){
        let returnObj = this.state.inTeams  ?
            <React.Fragment>
                <AuthenticatedTemplate>
                    {this.props.children}
                </AuthenticatedTemplate>
                <UnauthenticatedTemplate>                 
                    <div style={{ height: "calc(100vh - 72px)" }}>
                        <div className="authNotification" style={{
                            display: 'flex',
                            justifyContent: 'center',
                            alignItems: 'center',
                            width: '100%',
                            height: '100%'
                        }}>
                            <div className="authNotificationContents" style={{ textAlign: 'center' }}>
                                <img src='/images/freLauncher.svg' width='140px' /><br /><br />
                                <Text variant={'xxLarge'}>Signing you in</Text><br />
                                <Text variant='medium'>Just a few moments and we'll get you there</Text>
                            </div>
                        </div>
                    </div>
                </UnauthenticatedTemplate>
            </React.Fragment>
            :
            <React.Fragment>
                <MsalAuthenticationTemplate
                    interactionType={InteractionType.Redirect}
                    authenticationRequest={loginRequest}
                >
                    {this.props.children}
                </MsalAuthenticationTemplate>
            </React.Fragment>;
        return returnObj;
    }

    private updateState (){
        this.setState({authenticatedInTeams: true})
    }
    
}
