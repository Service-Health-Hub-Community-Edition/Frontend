import * as React from 'react';
import { loadTheme, createTheme } from '@fluentui/react';
import { NavAdmin } from "./NavAdmin";
import { AccessDenied } from "../AccessDenied";
import { acquireAccessToken } from "../../auth/AccessTokenHelper";
import { AuthenticationResult } from '@azure/msal-browser';
import { NavMenu } from '../NavMenu';
import { GlobalState } from '../GlobalState';
import { ThemeProvider, ITheme, MessageBar, MessageBarType } from '@fluentui/react';
import { IApplicationBrandingSettings } from "./applicationSettings/ApplicationConfig";
import './Admin.css';

interface IAdminState {
    width: number;
    height: number;
    documentHeight: number;
    error?: string;
    accessGranted?: boolean;
    theme?: any;
    accessToken?: AuthenticationResult;
    branding: IApplicationBrandingSettings;
    syncErrors: boolean;
}


export class AdminLayout extends React.Component<{}, IAdminState> {
    static contextType = GlobalState;

    constructor(props: {}) {
        super(props);

        this.state = {
            width: this._getWidth() - 64,
            height: this._getHeight() - 80,
            documentHeight: this._getDocumentHeight(),
            error: undefined,
            accessGranted: undefined,
            theme: undefined,
            accessToken: undefined,
            branding: {
                enabled: false,
                logo: '',
                backgroundColor: ''
            },
            syncErrors: false
        };

        window.addEventListener('resize', this._handleResize);
        document.addEventListener('readystatechange', this._handleResize);
    }

    public render() {
        const {
            error, width, height, accessGranted, theme, branding, syncErrors
        } = this.state;

        const documentHeight: number = this._getDocumentHeight();

        if (accessGranted === undefined) {
            return (<div />);
        } else {
            if (accessGranted === false) {
                return (
                    <div>
                        <AccessDenied />
                    </div>
                );
            }
        }

        const atJobsUrl: boolean = this._getUrlRelativePath().toLowerCase().startsWith('/admin/jobs');

        return (
            <ThemeProvider theme={theme} style={{ textDecoration: 'inherit' }} >
                <div className="shh-app">
                    <div className="shh-app-container">
                        <div className="shh-header-root" style={{
                            backgroundColor: branding.enabled && branding.backgroundColor !== undefined && branding.backgroundColor.trim() !== '' ? branding.backgroundColor.trim() : undefined
                        }} >
                            <div className="shh-header-nav-root" >
                                <NavMenu />
                            </div>
                        </div>
                        <div className="shh-content-root">
                            <div className="shh-content-container">
                                <div className="shh-content-menu" >
                                    <NavAdmin />
                                </div>
                                <div className="shh-content-main-content">
                                    {syncErrors && !atJobsUrl ? (
                                        <MessageBar
                                            messageBarType={MessageBarType.error}
                                            isMultiline={false}
                                        >
                                            At least one of the jobs failed. Click <a href='/admin/jobs'>here</a> to review.
                                        </MessageBar>) : (<></>)}
                                    {this.props.children}
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </ThemeProvider>
        );
    }

    componentDidMount(): void {
        let globalState: any = this.context;
        var theme: ITheme = globalState.getTheme();

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

        this.setState({
            theme: theme,
            branding: brandingObj!
        });

        this._onLoadAdmin();
    }

    private _onLoadAdmin(): void {
        const requiredRoles: string[] = ['Admin'];
        var userHasRequiredRole: boolean = false;

        acquireAccessToken()
            .then((response) => {
                var tokenClaims: any = response.account?.idTokenClaims;
                const userRoles: any = tokenClaims?.roles;
                userHasRequiredRole = userRoles.some((r: string) => requiredRoles.includes(r));

                fetch('/api/jobs/', { headers: response.idToken === "" ? {} : { 'Authorization': `Bearer ${response.idToken}` } })
                    .then(response => {
                        if (response.ok) {
                            return response.json();
                        } else return []
                    })
                    .then(response => {
                        var syncErrors: boolean = Array.isArray(response) ? response.filter((j: any) => j.state?.toString().trim().toLowerCase() === 'failed').length > 0 : false;

                        this.setState({
                            syncErrors: syncErrors
                        });
                    })

                this.setState({
                    accessGranted: userHasRequiredRole,
                    accessToken: response
                });

            });
    }

    _handleResize = () => {
        this.setState({
            height: this._getHeight() - 80,
            width: this._getWidth() - 64,
            documentHeight: this._getDocumentHeight()
        });
    }

    private _getHeight() {
        return window.innerHeight && document.documentElement.clientHeight ?
            Math.min(window.innerHeight, document.documentElement.clientHeight) :
            window.innerHeight ||
            document.documentElement.clientHeight ||
            document.getElementsByTagName('body')[0].clientHeight;
    }

    private _getDocumentHeight() {
        var body = document.body,
            html = document.documentElement;

        return Math.max(body.scrollHeight, body.offsetHeight,
            html.clientHeight, html.scrollHeight, html.offsetHeight);
    }

    private _getWidth() {
        return window.innerWidth && document.documentElement.clientWidth ?
            Math.min(window.innerWidth, document.documentElement.clientWidth) :
            window.innerWidth ||
            document.documentElement.clientWidth ||
            document.getElementsByTagName('body')[0].clientWidth;
    }

    private _getUrlRelativePath = (): string => {
        var url = document.location.toString();
        var arrUrl = url.split("//");

        var start = arrUrl[1].indexOf("/");
        var relUrl = arrUrl[1].substring(start);// "stop" is omitted, intercept all characters from start to end

        if (relUrl.indexOf("?") != -1) {
            relUrl = relUrl.split("?")[0];
        }
        return relUrl.toLowerCase();
    }
}