import * as React from 'react';
import { Toggle, Label, Text, Spinner, SpinnerSize } from '@fluentui/react';
import { acquireAccessToken } from "../../../auth/AccessTokenHelper";
import { AuthenticationResult } from '@azure/msal-browser';

interface IJobSettings {
    enabled: boolean;
}

interface IJobSettingsEditorState {
    componentId: string;
    settings: IJobSettings;
    accessGranted: boolean;
    initialized: boolean;
    error?: string;
}

export class JobSettingsEditor extends React.Component<{ componentId: string }, IJobSettingsEditorState> {

    constructor(props: { componentId: string }) {
        super(props);

        this.state = {
            componentId: props.componentId,
            settings: {
                enabled: false
            },
            accessGranted: false,
            initialized: false
        }
    }

    render() {
        const { settings, accessGranted, initialized, error } = this.state;

        if (!initialized)
            return (<>
                <br /><br />
                <Spinner size={SpinnerSize.medium} />
            </>);

        if (error !== undefined)
            return (<>
                <br /><br />
                <Text variant='medium'><b>{error}</b></Text>
            </>);

        if (!accessGranted)
            return (<>
                <br /><br />
                <Text variant='medium'><b>You have no access to this ressource.</b></Text>
            </>);

        return (<>
            <Label>Enabled</Label>
            <Toggle
                onText="On"
                offText="Off"
                checked={settings.enabled}
                onChange={(event: React.MouseEvent<HTMLElement>, checked?: boolean) => this._setConfigValue('enabled', checked ? checked : false)} />
        </>);
    }

    componentDidMount() {
        this.setState({
            componentId: this.props.componentId
        });

        this._loadData();
    }

    _loadData(): void {
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
                if (userHasRequiredRole)
                    fetch('/api/jobs/' + this.state.componentId + '/settings', { headers: authResponse.idToken === "" ? {} : { 'Authorization': `Bearer ${authResponse.idToken}` } })
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
                        this.setState({
                            settings: response,
                            initialized: true
                        });
                    })
            }).catch((err) => {
                this.setState({
                    error: err.message
                });
            });
    }

    _saveData(settings: IJobSettings): void {
        var params: RequestInit;

        acquireAccessToken()
            .then((response) => {
                params = {
                    headers: {
                        "Content-Type": "application/json charset=UTF-8",
                        "Authorization": "Bearer " + response.idToken
                    },
                    body: JSON.stringify(settings),
                    method: "POST"
                };

                const uri: string = '/api/jobs/' + this.state.componentId + '/settings';

                fetch(uri, params)
                    .then((response) => {
                        if (!response.ok) {                            
                            this.setState({
                                error: "Couldn't save job settings. Error code: " + response.status
                            });
                        }
                    });
            });
    }

    private _setConfigValue<Key extends keyof IJobSettings>(key: Key, value: IJobSettings[Key]): void {
        var settings: IJobSettings = this.state.settings ?
            this.state.settings :
            {
                enabled: false
            }

        settings[key] = value;

        this._saveData(settings);

        this.setState({
            settings: settings
        });
    }
}