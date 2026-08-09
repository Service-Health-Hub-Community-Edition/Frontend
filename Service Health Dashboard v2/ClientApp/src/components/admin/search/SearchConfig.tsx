import * as React from 'react';
import {
    SelectionMode, IObjectWithKey, ICommandBarProps,
    Panel, PanelType, Dialog, DialogType, TextField, Link, ILinkStyleProps, ILinkStyles,
    PrimaryButton, DefaultButton, DialogFooter, Label, Toggle, Spinner, SpinnerSize, Dropdown, IDropdownOption, ISelectableOption, Text, TextStyles,
    Separator, SpinButton, CommandBar, ICommandBarItemProps, Icon, IconButton
} from '@fluentui/react';
import { DetailPageHeader } from '@m365-admin/detail-page';
import { IconAlert, IconAlertSize, IconAlertStatus } from '@m365-admin/icon-alert';
import { CompositeList, ICompositeListColumn, ICompositeListSelectionMap, ICompositeListSelectionMapItem } from '@m365-admin/composite-list';
import { EmptyStateImageSize } from '@m365-admin/empty-state';
import { M365Breadcrumb } from '@m365-admin/m365-breadcrumb';
import { AccessDenied } from "../../AccessDenied";
import { acquireAccessToken } from "../../../auth/AccessTokenHelper";
import { AuthenticationResult } from '@azure/msal-browser';
import { AdminLayout } from '../AdminLayout';
import { stringify } from 'uuid';

interface ISearchConnectorDetails {
    schemaUpdateAvailable: boolean;
    connectorStatus: string;
}
interface ISearchConnectorSettings {
    configured: boolean;
    enabled: boolean;
    schemaUpdated: Date;
    rootUrl: string;
}

interface ISearchConfig {
    connectorSettings: ISearchConnectorSettings;
    graphAPIConnector: any;
    details: ISearchConnectorDetails;
}

interface ISearchConfigState {
    persistedSettings?: ISearchConfig;
    settings?: ISearchConfig;
    rootUrlEditMode: boolean;
    newRootUrl: string;
    initialized: boolean;
    accessGranted: boolean;
    error?: string;
}

const labelWidth: string = '300px';
const groupMaxWidth: string = '750px';

export class SearchConfig extends React.Component<{}, ISearchConfigState> {

    constructor(props: {}) {
        super(props);

        this.state = {
            persistedSettings: undefined,
            settings: undefined,
            initialized: false,
            rootUrlEditMode: false,
            newRootUrl: "",
            accessGranted: false,
            error: undefined
        };
    }

    public render() {
        const {
            settings, initialized, rootUrlEditMode, newRootUrl
        } = this.state;

        if (!initialized)
            return (<div className="container" style={{ height: "calc(100vh - 72px)" }}>
                <div className="row" style={{
                    display: 'flex',
                    justifyContent: 'center',
                    alignItems: 'center',
                    width: '100%',
                    height: '100%'
                }}>
                    <div className="col">
                        <Spinner label="Loading data..." size={SpinnerSize.large} />
                    </div>
                </div>
            </div>);

        return (
            <AdminLayout>
                <div className="container" style={{ maxWidth: '97%' }}>
                    <div className="row">
                        <div className="col">
                            <M365Breadcrumb
                                items={[
                                    { text: 'Home', key: 'home', href: '/admin' },
                                    { text: 'Search', key: 'search', isCurrentItem: true }
                                ]}
                                style={{ marginBottom: '16px' }}
                            />
                            <DetailPageHeader
                                title="Search settings (preview)"
                                description="Set up and configure the Service Health Hub search settings. Create and configure Graph Connector to index service communication data in Microsoft Search."

                            />
                        </div>
                    </div>
                    <div className="row" style={{ marginBottom: '40px', maxWidth: groupMaxWidth }} >
                        <div className="col">
                            <Text variant='large'><b>Status</b></Text>
                            <Separator />
                            <Text variant='medium'>Service Health Hub Graph connector status.</Text><br />


                            <div className="container" style={{ marginTop: '24px', display: 'inline-table', justifyContent: 'flex-start', justifySelf: 'flex-start' }} >
                                <div className="row" style={{ marginBottom: '12px' }}>
                                    <div style={{ width: labelWidth }} >
                                        <Label>Connected</Label>
                                    </div>
                                    <div className="col" >
                                        <Text variant="medium">{settings?.connectorSettings.configured ? "Yes" : "No"}</Text>
                                    </div>
                                </div>
                                <div className="row" style={{ marginBottom: '12px' }}>
                                    <div style={{ width: labelWidth }} >
                                        <Label>Enabled</Label>
                                    </div>
                                    <div className="col" >
                                        <Text variant="medium">{settings?.connectorSettings.enabled ? "Yes" : "No"}</Text>
                                    </div>
                                </div>
                                <div className="row" style={{ marginBottom: '12px' }}>
                                    <div style={{ width: labelWidth }} >
                                        <Label>Connector state</Label>
                                    </div>
                                    <div className="col" >
                                        <Text variant="medium">{settings?.details.connectorStatus}</Text>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>

                    <div className="row" style={{ marginBottom: '40px', maxWidth: groupMaxWidth }} >
                        <div className="col">
                            <Text variant='large'><b>Service Health Hub</b></Text>
                            <Separator />
                            <Text variant='medium'>Service Health Hub search settings. Updating root URL will trigger a full index process.</Text><br />

                            <div className="container" style={{ marginTop: '24px', display: 'inline-table', justifyContent: 'flex-start', justifySelf: 'flex-start' }} >
                                <div className="row" style={{ marginBottom: '12px' }} >
                                    <div style={{ width: labelWidth }} >
                                        <Label>Root URL</Label>
                                    </div>
                                    <div className="col" style={{ paddingTop: '6px' }} >
                                        {rootUrlEditMode ?
                                            (
                                                <>
                                                    <TextField
                                                        defaultValue={settings?.connectorSettings.rootUrl}
                                                        onChange={(event: any, newValue: string|undefined) => this._updateRootUrlCachedValue(newValue)}
                                                    />
                                                    <IconButton
                                                        iconProps={{ iconName: 'Accept', color: 'green' }}
                                                        title="Update"
                                                        ariaLabel="Update"
                                                        disabled={newRootUrl === "" || newRootUrl.toLowerCase().trim() === settings?.connectorSettings.rootUrl.toLowerCase().trim()}
                                                        onClick={() => this._updateRootUrl()}
                                                    />
                                                    <IconButton
                                                        iconProps={{ iconName: 'Cancel', color: 'red' }}
                                                        title="Cancel"
                                                        ariaLabel="Cancel"
                                                        onClick={() => this._cancelEditRootUrl()} />
                                                </>
                                            ) :
                                            (
                                                <>
                                                    <Text>
                                                        {settings?.connectorSettings.rootUrl ? settings?.connectorSettings.rootUrl : ""}
                                                    </Text>
                                                    <IconButton iconProps={{ iconName: 'Edit' }} title="Edit" ariaLabel="Edit" onClick={() => this._editRootUrl()} />
                                                </>
                                            )
                                        }
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>

                    <div className="row" style={{ maxWidth: groupMaxWidth, paddingBottom: '40px' }}>
                        {settings?.connectorSettings.configured ?
                            (
                                settings?.connectorSettings.enabled ? (
                                    <div className="col">
                                        <PrimaryButton onClick={() => this._onDisableConnector()}>
                                            Disable
                                        </PrimaryButton>
                                    </div>
                                ) : (
                                    <div className="col">
                                        <PrimaryButton onClick={() => this._onEnableConnector()}>
                                            Enable
                                        </PrimaryButton>
                                    </div>
                                )
                            ) : (
                                <div className="col">
                                    <PrimaryButton onClick={() => this._onConnectConnector()}>
                                        Connect
                                    </PrimaryButton>
                                </div>
                            )}
                        {settings?.connectorSettings.configured && settings?.details.connectorStatus.toLowerCase() === "draft" ?
                            (
                                <div className="col">
                                    <PrimaryButton onClick={() => this._onUpdateSchema()}>
                                        Configure schema
                                    </PrimaryButton>
                                </div>
                            ) :
                            (<div />)}
                    </div>
                </div>
            </AdminLayout>
        );
    }

    componentDidMount() {
        this._onLoadConfig();
    }

    private _onLoadConfig(): void {
        const requiredRoles: string[] = ['Admin'];
        var authResponse: AuthenticationResult;
        var userHasRequiredRole: boolean = false;

        this.setState({
            initialized: false
        });

        acquireAccessToken()
            .then((response) => {
                var tokenClaims: any = response.account?.idTokenClaims;
                const userRoles: any = tokenClaims?.roles;
                userHasRequiredRole = userRoles.some((r: string) => requiredRoles.includes(r));

                this.setState({
                    accessGranted: userHasRequiredRole
                });

                authResponse = response;
            })
            .then(() => {
                if (userHasRequiredRole) fetch('/api/admin/search', { headers: authResponse.idToken === "" ? {} : { 'Authorization': `Bearer ${authResponse.idToken}` } })
                    .then(response => {
                        if (response.ok) {
                            return response.json();
                        } else {
                            this.setState({
                                error: response.status + " " + response.statusText
                            });
                            throw Error(response.status + " " + response.statusText);
                        }
                    })
                    .then(response => {
                        const persistedSettings: ISearchConfig = JSON.parse(JSON.stringify(response));
                        this.setState({
                            settings: response,
                            persistedSettings: persistedSettings,
                            initialized: true
                        });
                    })
            }).catch((err) => {
                this.setState({
                    error: err.message
                });
            });
    }

    private _onEnableConnector(): void {
        var request: any = {
            operation: "enable"
        }

        this._processPostRequest(JSON.stringify(request));
    }

    private _onDisableConnector(): void {
        var request: any = {
            operation: "disable"
        }

        this._processPostRequest(JSON.stringify(request));
    }

    private _onConnectConnector(): void {
        var request: any = {
            operation: "create",
            rootUrl: ""
        }

        this._processPostRequest(JSON.stringify(request));
    }

    private _onUpdateSchema(): void {
        var request: any = {
            operation: "schemaUpdate"
        }

        this._processPostRequest(JSON.stringify(request));
    }

    _processPostRequest(body: string): void {
        var params: RequestInit;
        acquireAccessToken()
            .then((response) => {
                params = {
                    headers: {
                        "Content-Type": "application/json charset=UTF-8",
                        "Authorization": "Bearer " + response.idToken
                    },
                    body: body,
                    method: "POST"
                };

                const uri: string = '/api/admin/search';

                fetch(uri, params)
                    .then((response) => {
                        if (!response.ok) {
                            // make the promise be rejected if we didn't get a 2xx response
                            const err = new Error("Couldn't process request. Error details:<br/><br/><b>HTTP " + response.status + "</b><br/>" + response.statusText);
                            throw err;
                        } else {
                            return response.json();
                        }
                    })
                    .then((response) => {

                        const persistedSettings: ISearchConfig = JSON.parse(JSON.stringify(response));
                        this.setState({
                            settings: response,
                            persistedSettings: persistedSettings,
                            initialized: true
                        });
                    });
            });
    }

    private _editRootUrl() {
        this.setState({
            rootUrlEditMode: true
        });
    }

    private _cancelEditRootUrl() {
        this.setState({
            rootUrlEditMode: false,
            newRootUrl: ""
        });
    }

    private _updateRootUrlCachedValue(value: string | undefined): void {
        var newValue: string = value === undefined ? "" : value
        
        this.setState({
            newRootUrl: newValue
        });
    }

    private _updateRootUrl() {
        if (this.state.newRootUrl.trim() !== "") {
            var request: any = {
                operation: "create",
                rootUrl: this.state.newRootUrl
            }

            this._processPostRequest(JSON.stringify(request));

            this.setState({
                rootUrlEditMode: false,
                newRootUrl: ""
            });
        }
    }
}