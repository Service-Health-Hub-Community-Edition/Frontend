import * as React from 'react';
import { Component } from 'react';
import { checkInTeams } from '../auth/detectTeams';
import { NavMenu } from '../NavMenu';
import { NavAzure } from './NavAzure';
import { MessageBar, MessageBarType } from '@fluentui/react';
import { IApplicationBrandingSettings } from "../admin/applicationSettings/ApplicationConfig";
import { AuthenticationResult } from '@azure/msal-browser';
import { acquireAccessToken } from "../../auth/AccessTokenHelper";

interface ILayoutProps {
    hideMenu: boolean;
    inTeams: boolean;
    branding: IApplicationBrandingSettings;
    syncErrors: boolean;
}

export class AzureLayout extends Component<{ hideMenu?: boolean }, ILayoutProps> {
    static displayName = AzureLayout.name;

    constructor(props: { hideMenu?: boolean }) {
        super(props);
 
        this.state = {
            hideMenu: props.hideMenu ? props.hideMenu : false,
            inTeams: checkInTeams(),
            branding: {
                enabled: false,
                logo: '',
                backgroundColor: ''
            },
            syncErrors: false
        }
    }

    render() {
        const { hideMenu, branding, syncErrors } = this.state;

        return (
            <div className="shh-app">
                <div className="shh-app-container">
                    {
                        this.state.inTeams ? "" : (
                            <div className="shh-header-root" style={{
                                backgroundColor: branding.enabled && branding.backgroundColor !== undefined && branding.backgroundColor.trim() !== '' ? branding.backgroundColor.trim() : undefined
                            }} >
                                <div className="shh-header-nav-root" >
                                    <NavMenu />
                                </div>
                            </div>)
                    }

                    <div className="shh-content-root">
                        <div className="shh-content-container">

                            {hideMenu ?
                                "" :
                                (<div className="shh-content-menu" >
                                    <NavAzure />
                                </div>)}

                            <div className="shh-content-main-content">
                                {syncErrors ? (
                                    <MessageBar
                                        messageBarType={MessageBarType.error}
                                        isMultiline={false}
                                    >
                                        At least one of the jobs failed. {this.state.inTeams ? "Review job statistics in Service Health Hub Admin Center" : (<>Click <a href='/admin/jobs'>here</a> to review</>)}.
                                    </MessageBar>) : (<></>)}
                                {this.props.children}
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        );
    }

    componentDidMount() {
        var branding = sessionStorage.getItem('shh_branding');
        var brandingObj: IApplicationBrandingSettings | undefined = undefined;

        if (branding !== null && branding !== undefined && branding.trim() !== '')
            try {
                brandingObj = JSON.parse(branding);
            } catch {
                brandingObj = {
                    enabled: false,
                    logo: '',
                    backgroundColor: ''
                }
            }

        const requiredRoles: string[] = ['Admin'];
        var authResponse: AuthenticationResult;
        var userHasRequiredRole: boolean = false;

        acquireAccessToken()
            .then((response: any) => {
                var tokenClaims: any = response.account?.idTokenClaims;
                const userRoles: any = tokenClaims?.roles;
                userHasRequiredRole = userRoles.some((r: string) => requiredRoles.includes(r));

                authResponse = response;
            })
            .then(() => {
                if (userHasRequiredRole) {
                    fetch('/api/jobs/', { headers: authResponse.idToken === "" ? {} : { 'Authorization': `Bearer ${authResponse.idToken}` } })
                        .then(response => {
                            if (response.ok) {
                                return response.json();
                            } else return []
                        })
                        .then(response => {
                            var syncErrors: boolean = Array.isArray(response) ? response.filter((j: any) => j.state?.toString().trim().toLowerCase() === 'failed').length > 0 : false;

                            this.setState({
                                branding: brandingObj!,
                                syncErrors: syncErrors
                            });
                        })
                } else {
                    this.setState({
                        branding: brandingObj!
                    });
                }
            }).catch((err) => {
                this.setState({
                    branding: brandingObj!
                });
            });
    }
}

