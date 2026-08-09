import * as React from 'react';
import {
    SelectionMode, IObjectWithKey, ICommandBarProps,
    Panel, PanelType, Dialog, DialogType, TextField, Link, ILinkStyleProps, ILinkStyles,
    PrimaryButton, DefaultButton, DialogFooter, Label, Toggle, Spinner, SpinnerSize, Dropdown, IDropdownOption, ISelectableOption, Text, TextStyles,
    Separator, SpinButton, CommandBar, ICommandBarItemProps
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
import { clientVersion } from '../../../Version';

export interface IApplicationBrandingSettings {
    enabled?: boolean;
    backgroundColor?: string;
    logo?: string;
}

interface IApplicationSettings {
    pirStorageUri: string;
    releaseMessageEndpoint: string;
    azureTranslatorEnabled: boolean;
    azureTranslatorLocation: string;
    azureTranslatorKey: string;
    languageServiceEnabled: boolean;
    languageServiceEndpoint: string;
    languageServiceKey: string;
    textSummarizationSentenceCount: number;
    branding: IApplicationBrandingSettings;
}

interface IApplicationConfigState {
    persistedSettings?: IApplicationSettings;
    settings?: IApplicationSettings;
    initialized: boolean;
    accessGranted: boolean;
    error?: string;
}

const labelWidth: string = '300px';
const groupMaxWidth: string = '750px';

export class ApplicationConfig extends React.Component<{}, IApplicationConfigState> {
    static dbVersion: string | null = sessionStorage.getItem("shh_dbVersion");
    logoFileSelector: HTMLInputElement | null = null;

    constructor(props: {}) {
        super(props);

        this.state = {
            persistedSettings: undefined,
            settings: undefined,
            initialized: false,
            accessGranted: false,
            error: undefined
        };
    }

    public render() {
        const {
            settings, initialized
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

        var logoCommandItems: ICommandBarItemProps[] = [
            {
                key: 'logoUpload',
                text: 'Upload',
                iconProps: { iconName: 'Upload' },
                onClick: () => { this._openFileSelectDialog() }
            }];

        if (settings?.branding?.logo && settings?.branding?.logo.trim() !== '')
            logoCommandItems.push({
                key: 'logoRemove',
                text: 'Remove',
                iconProps: { iconName: 'Cancel' },
                onClick: () => { this._removeLogo() }
            });

        return (
            <AdminLayout>
            <div className="container" style={{ maxWidth: '97%' }}>
                <div className="row">
                    <div className="col">
                        <M365Breadcrumb
                            items={[
                                { text: 'Home', key: 'home', href: '/admin' },
                                { text: 'Settings', key: 'settings', isCurrentItem: true }
                            ]}
                            style={{ marginBottom: '16px' }}
                        />
                        <DetailPageHeader
                            title="Application settings"
                                description="Set up and configure the Service Health Hub application settings. Settings specified here will apply to the web app and synchronization components."
                                
                        />
                    </div>
                    </div>
                    <div className="row" style={{ marginBottom: '40px', maxWidth: groupMaxWidth }} >
                        <div className="col">
                            <Text variant='large'><b>Version</b></Text>
                            <Separator />
                            <Text variant='medium'>Service Health Hub version information. Contains version information for client application and database schema version.</Text><br />
                        

                            <div className="container" style={{ marginTop: '24px', display: 'inline-table', justifyContent: 'flex-start', justifySelf: 'flex-start' }} >
                                <div className="row" style={{ marginBottom: '12px' }}>
                                    <div style={{ width: labelWidth }} >
                                        <Label>Client version</Label>
                                    </div>
                                    <div className="col" >
                                        <Text variant="medium">{clientVersion}</Text>
                                    </div>
                                </div>
                                <div className="row" style={{ marginBottom: '12px' }}>
                                    <div style={{ width: labelWidth }} >
                                        <Label>Database schema</Label>
                                    </div>
                                    <div className="col" >
                                        <Text variant="medium">{ApplicationConfig.dbVersion}</Text>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>

                    <div className="row" style={{ marginBottom: '40px', maxWidth: groupMaxWidth }} >
                        <div className="col">
                            <Text variant='large'><b>Branding</b></Text>
                            <Separator />
                            <Text variant='medium'>Use your organization's logo and custom background color to provide a consistent look-and-feel. Organizational logo must be transparent PNG and can be any size, however it will be resized. For the best quality and performance, please provide a PNG file with height of 24 pixels.</Text><br />

                            <div className="container" style={{ marginTop: '24px', display: 'inline-table', justifyContent: 'flex-start', justifySelf: 'flex-start' }} >
                                <div className="row" style={{ marginBottom: '12px' }} >
                                    <div style={{ width: labelWidth }} >
                                        <Label>Branding</Label>
                                    </div>
                                    <div className="col" style={{ paddingTop: '6px' }} >
                                        <Toggle
                                            onText="On"
                                            offText="Off"
                                            checked={settings?.branding?.enabled}
                                            onChange={(event: React.MouseEvent<HTMLElement>, checked?: boolean) => this.setBrandingConfigValue('enabled', checked ? checked : false)} />
                                    </div>
                                </div>
                                <div className="row" style={{ marginBottom: '12px' }}>
                                    <div style={{ width: labelWidth }} >
                                        <Label
                                            disabled={!settings?.branding?.enabled}>
                                            Navigation bar color
                                        </Label>
                                    </div>
                                    <div className="col" >
                                        {settings?.branding?.enabled ? (
                                            <TextField
                                                disabled={!settings?.branding?.enabled}
                                                value={settings?.branding?.backgroundColor}
                                                onGetErrorMessage={this.getErrorMessage}
                                                validateOnFocusOut
                                                onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => this.setBrandingConfigValue('backgroundColor', newValue ? newValue : "")}
                                            />) : (
                                            <TextField
                                                disabled={!settings?.branding?.enabled}
                                                value={settings?.branding?.backgroundColor}
                                            />
                                        )}
                                    </div>
                                </div>
                                <div className="row" style={{ marginBottom: '12px' }}>
                                    <div style={{ width: labelWidth }} >
                                        <Label
                                            disabled={!settings?.branding?.enabled}>
                                            Logo
                                        </Label>
                                    </div>
                                    <div className="col" >
                                        <input
                                            type='file'
                                            id='logoFile'
                                            accept='image/png'
                                            ref={(ref) => this.logoFileSelector = ref}
                                            style={{ display: 'none' }}
                                            onChange={(ev: React.ChangeEvent<HTMLInputElement>) => this._uploadFile(ev)}
                                        />
                                        {settings?.branding?.enabled ? (<>
                                            <CommandBar items={logoCommandItems} styles={{ root: { padding: '0px' } }} /><br />
                                            
                                            {settings?.branding?.logo && settings?.branding?.logo.trim() !== '' ? (
                                                <div style={{
                                                    height: '32px',
                                                    backgroundColor: '#808080',
                                                    display: 'flex',
                                                    flexWrap: 'wrap',
                                                    justifyContent: 'center',
                                                    alignContent: 'center'
                                                }} >
                                                    <img src={settings?.branding?.logo} height='24px' />
                                                </div>) : ""}
                                        </>) : settings?.branding?.logo && settings?.branding?.logo.trim() !== '' ? (
                                                <div style={{
                                                    height: '32px',
                                                    backgroundColor: '#808080',
                                                    display: 'flex',
                                                    flexWrap: 'wrap',
                                                    justifyContent: 'center',
                                                    alignContent: 'center',
                                                    opacity: '0.3'
                                                }} >
                                                    <img src={settings?.branding?.logo} height='24px' style={{ opacity: '0.5' }} />
                                                </div>
                                        ) : ""}
                                        
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>

                    <div className="row" style={{ marginBottom: '40px', maxWidth: groupMaxWidth }} >
                    <div className="col">
                            <Text variant='large'><b>Azure Translation</b></Text>
                            <Separator />
                            <Text variant='medium'>Service Health Hub can automatically translate content to languages specified in notification routing rules and user settings. You can set up and manage translation options here.</Text><br />
                            <Link href='https://learn.microsoft.com/en-us/azure/cognitive-services/translator/how-to-create-translator-resource#get-your-authentication-keys-and-endpoint' target='_blank'>
                                Read more
                            </Link>

                            <div className="container" style={{ marginTop: '24px', display: 'inline-table', justifyContent: 'flex-start', justifySelf: 'flex-start' }} >
                                <div className="row" style={{ marginBottom: '12px' }} >
                                    <div style={{ width: labelWidth }} >
                                        <Label>Azure Translation</Label>
                                    </div>
                                    <div className="col" style={{ paddingTop: '6px' }} >
                                        <Toggle
                                            onText="On"
                                            offText="Off"
                                            checked={settings?.azureTranslatorEnabled}
                                            onChange={(event: React.MouseEvent<HTMLElement>, checked?: boolean) => this.setConfigValue('azureTranslatorEnabled', checked ? checked : false)} />
                                    </div>
                                </div>
                                <div className="row" style={{ marginBottom: '12px' }}>
                                    <div style={{ width: labelWidth }} >
                                        <Label
                                            disabled={!settings?.azureTranslatorEnabled}>
                                            Resource location
                                        </Label>
                                    </div>
                                    <div className="col" >
                                        {settings?.azureTranslatorEnabled ? (
                                            <TextField
                                                disabled={!settings?.azureTranslatorEnabled}
                                                value={settings?.azureTranslatorLocation}
                                                required
                                                onGetErrorMessage={this.getErrorMessage}
                                                validateOnFocusOut
                                                onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => this.setConfigValue('azureTranslatorLocation', newValue ? newValue : "")}
                                            />) : (
                                            <TextField
                                                disabled={!settings?.azureTranslatorEnabled}
                                                value={settings?.azureTranslatorLocation}
                                            />
                                        )}
                                    </div>
                                </div>
                                <div className="row" style={{ marginBottom: '12px' }}>
                                    <div style={{ width: labelWidth }} >
                                        <Label
                                            disabled={!settings?.azureTranslatorEnabled}>
                                            Subscription key
                                        </Label>
                                    </div>
                                    <div className="col" >
                                        {settings?.azureTranslatorEnabled ? (
                                            <TextField
                                                type="password"
                                                disabled={!settings?.azureTranslatorEnabled}
                                                value={settings?.azureTranslatorKey}
                                                required
                                                onGetErrorMessage={this.getErrorMessage}
                                                validateOnFocusOut
                                                onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => this.setConfigValue('azureTranslatorKey', newValue ? newValue : "")}
                                            />) : (
                                            <TextField
                                                type="password"
                                                disabled={!settings?.azureTranslatorEnabled}
                                                value={settings?.azureTranslatorKey}
                                            />
                                        )}
                                    </div>
                                </div>
                            </div>
                    </div>
                    </div>

                    <div className="row" style={{ marginBottom: '40px', maxWidth: groupMaxWidth }} >
                        <div className="col">
                            <Text variant='large'><b>Language service</b></Text>
                            <Separator />
                            <Text variant='medium'>Service Health Hub can create a short a summary of the communication during the synchronization. The summary is visible as a tooltip in the lists and on top of communication details, in the details panel. You can set up and manage language service options here.</Text><br />
                            <Link href='https://aka.ms/languageservice' target='_blank'>
                                Read more
                            </Link>

                            <div className="container" style={{ marginTop: '24px', display: 'inline-table', justifyContent: 'flex-start', justifySelf: 'flex-start' }} >
                                <div className="row" style={{ marginBottom: '12px' }} >
                                    <div style={{ width: labelWidth }} >
                                        <Label>Language service</Label>
                                    </div>
                                    <div className="col" style={{ paddingTop: '6px' }} >
                                        <Toggle
                                            onText="On"
                                            offText="Off"
                                            checked={settings?.languageServiceEnabled}
                                            onChange={(event: React.MouseEvent<HTMLElement>, checked?: boolean) => this.setConfigValue('languageServiceEnabled', checked ? checked : false)} />
                                    </div>
                                </div>
                                <div className="row" style={{ marginBottom: '12px' }}>
                                    <div style={{ width: labelWidth }} >
                                        <Label
                                            disabled={!settings?.languageServiceEnabled}>
                                            Endpoint
                                        </Label>
                                    </div>
                                    <div className="col" >
                                        {settings?.languageServiceEnabled ? (
                                            <TextField
                                                disabled={!settings?.languageServiceEnabled}
                                                value={settings?.languageServiceEndpoint}
                                                required
                                                onGetErrorMessage={this.getErrorMessage}
                                                validateOnFocusOut
                                                onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => this.setConfigValue('languageServiceEndpoint', newValue ? newValue : "")}
                                            />) : (
                                            <TextField
                                                    disabled={!settings?.languageServiceEnabled}
                                                    value={settings?.languageServiceEndpoint}
                                            />
                                        )}
                                    </div>
                                </div>
                                <div className="row" style={{ marginBottom: '12px' }}>
                                    <div style={{ width: labelWidth }} >
                                        <Label
                                            disabled={!settings?.languageServiceEnabled}>
                                            Subscription key
                                        </Label>
                                    </div>
                                    <div className="col" >
                                        {settings?.languageServiceEnabled ? (
                                            <TextField
                                                type="password"
                                                disabled={!settings?.languageServiceEnabled}
                                                value={settings?.languageServiceKey}
                                                required
                                                onGetErrorMessage={this.getErrorMessage}
                                                validateOnFocusOut
                                                onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => this.setConfigValue('languageServiceKey', newValue ? newValue : "")}
                                            />) : (
                                            <TextField
                                                type="password"
                                                    disabled={!settings?.languageServiceEnabled}
                                                    value={settings?.languageServiceKey}
                                            />
                                        )}
                                    </div>
                                </div>
                                <div className="row" style={{ marginBottom: '12px' }}>
                                    <div style={{ width: labelWidth }} >
                                        <Label
                                            disabled={!settings?.languageServiceEnabled}>
                                            Sentence count
                                        </Label>
                                    </div>
                                    <div className="col" >
                                        <SpinButton
                                            disabled={!settings?.languageServiceEnabled}
                                            value={settings?.textSummarizationSentenceCount.toString()}
                                            min={2}
                                            max={5}
                                            step={1}
                                            incrementButtonAriaLabel="Increase value by 1"
                                            decrementButtonAriaLabel="Decrease value by 1"
                                            styles={{ spinButtonWrapper: { width: 50 } }}
                                            onChange={(event: React.SyntheticEvent<HTMLElement>, newValue?: string) => {
                                                const numericValue = newValue ? Number(newValue) : 2;
                                                this.setConfigValue('textSummarizationSentenceCount', numericValue)
                                            }}
                                        />
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                    <div className="row" style={{ marginBottom: '40px', maxWidth: groupMaxWidth }}>
                        <div className="col">
                            <Text variant='large'><b>Service Health Hub Release manager</b></Text>
                            <Separator />
                            <Text variant='medium'>Configure Service Health Hub release manager endpoint. Release manager endpoint provides you with the information about new features and planned releases.</Text>

                            <div className="container" style={{ marginTop: '24px', display: 'inline-table', justifyContent: 'flex-start', justifySelf: 'flex-start' }} >
                                <div className="row" style={{ marginBottom: '12px' }}>
                                    <div style={{ width: labelWidth }} >
                                        <Label>Release manager endpoint</Label>
                                    </div>
                                    <div className="col" >
                                        <TextField
                                            value={settings?.releaseMessageEndpoint}
                                            required
                                            onGetErrorMessage={this.getErrorMessage}
                                            validateOnFocusOut
                                            onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => this.setConfigValue('releaseMessageEndpoint', newValue ? newValue : "")}
                                        />
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                    <div className="row" style={{ marginBottom: '40px', maxWidth: groupMaxWidth }}>
                        <div className="col">
                            <Text variant='large'><b>Post-incident review storage</b></Text>
                            <Separator />
                            <Text variant='medium'>Configure link to post-incident review blob storage. Specify an Azure Blob Storage Container link where post incident reports will be stored. Ensure that Service Health Hub app registration principal has Storage Blob Data Contributor role assigned on the container via Azure RBAC</Text><br />
                            <Link href='https://learn.microsoft.com/en-us/azure/storage/blobs/assign-azure-role-data-access?tabs=portal' target='_blank'>Read more</Link>

                            <div className="container" style={{ marginTop: '24px', display: 'inline-table', justifyContent: 'flex-start', justifySelf: 'flex-start' }} >
                                <div className="row" style={{ marginBottom: '12px' }}>
                                    <div style={{ width: labelWidth }} >
                                        <Label>Storage container URL</Label>
                                    </div>
                                    <div className="col" >
                                        <TextField
                                            value={settings?.pirStorageUri}
                                            onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => this.setConfigValue('pirStorageUri', newValue ? newValue : "")}
                                        />
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                    <div className="row" style={{ maxWidth: groupMaxWidth, paddingBottom: '40px' }}>
                        <div className="col">
                            <PrimaryButton disabled={!this.formValid()} onClick={() => this._onSaveConfig()}>
                                Save
                            </PrimaryButton>
                        </div>
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
                if (userHasRequiredRole) fetch('/api/settings/application', { headers: authResponse.idToken === "" ? {} : { 'Authorization': `Bearer ${authResponse.idToken}` } })
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
                        const persistedSettings: IApplicationSettings = JSON.parse(JSON.stringify(response));
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

    private _onSaveConfig(): void {
        var params: RequestInit;

        acquireAccessToken()
            .then((response) => {
                params = {
                    headers: {
                        "Content-Type": "application/json charset=UTF-8",
                        "Authorization": "Bearer " + response.idToken
                    },
                    body: JSON.stringify(this.state.settings),
                    method: "POST"
                };

                const uri: string = '/api/settings/application';

                fetch(uri, params)
                    .then((response) => {
                        if (!response.ok) {
                            // make the promise be rejected if we didn't get a 2xx response
                            const err = new Error("Couldn't save release. Error details:<br/><br/><b>HTTP " + response.status + "</b><br/>" + response.statusText);
                            throw err;
                        }

                        this.setState({
                            persistedSettings: JSON.parse(JSON.stringify(this.state.settings))
                        });

                        sessionStorage.setItem("shh_branding", this.state.settings?.branding ? JSON.stringify(this.state.settings?.branding) : "");
                    });
            });
    }

    private _openFileSelectDialog(): void {
        if (this.logoFileSelector && this.logoFileSelector !== null)
            this.logoFileSelector.click();
    }

    private _uploadFile(ev: React.ChangeEvent<HTMLInputElement>): void {
        const imageMimeType = /image\/(png)/i;
        var fileList: FileList | null = ev.target.files;
        if (fileList !== null && fileList.length > 0) {
            var file = fileList[0];
            if (!file.type.match(imageMimeType)) {
                alert('Image type is not valid. Please upload a valid PNG file.');
                return;
            }

            const fileReader = new FileReader();
            fileReader.onload = (e: ProgressEvent<FileReader>) => {
                if (e.target && e.target !== null) {
                    var result = e.target.result;
                    if (result) {
                        this.setBrandingConfigValue('logo', result.toString());
                    }
                }
            }

            fileReader.readAsDataURL(file);
        }
    }

    private _removeLogo(): void {
        this.setBrandingConfigValue('logo', '');
    }

    private setConfigValue<Key extends keyof IApplicationSettings>(key: Key, value: IApplicationSettings[Key]): void {
        var settings: IApplicationSettings = this.state.settings ?
            this.state.settings :
            {
                azureTranslatorEnabled: false,
                azureTranslatorKey: "",
                azureTranslatorLocation: "",
                pirStorageUri: "",
                releaseMessageEndpoint: "",
                languageServiceEnabled: false,
                languageServiceEndpoint: "",
                languageServiceKey: "",
                textSummarizationSentenceCount: 3,
                branding: {
                    enabled: false,
                    backgroundColor: '',
                    logo: ''
                }
            }

        settings[key] = value;

        this.setState({
            settings: settings
        });
    }

    private setBrandingConfigValue<Key extends keyof IApplicationBrandingSettings>(key: Key, value: IApplicationBrandingSettings[Key]): void {
        var settings: IApplicationBrandingSettings = this.state.settings?.branding ?
            this.state.settings.branding :
            {
                enabled: false,
                backgroundColor: '',
                logo: ''
            }

        settings[key] = value;

        this.setConfigValue('branding', settings);
    }

    private getErrorMessage(value: string): string {
        return value.length > 0 ? '' : 'Required value.';
    };

    private isNullOrWhiteSpace(value: string | undefined | null): boolean {
        return value === undefined ||
            value === null ||
            value.trim() === "";
    }

    private formValid(): boolean {
        if (this.state.settings === undefined)
            return false;

        const contentsValid: boolean = this.state.settings !== undefined &&
            (this.state.settings.azureTranslatorEnabled ?
            !this.isNullOrWhiteSpace(this.state.settings.azureTranslatorLocation) &&
            !this.isNullOrWhiteSpace(this.state.settings.azureTranslatorKey) :
                true) &&
            (this.state.settings.languageServiceEnabled ?
                !this.isNullOrWhiteSpace(this.state.settings.languageServiceEndpoint) &&
                !this.isNullOrWhiteSpace(this.state.settings.languageServiceKey) &&
                this.state.settings.textSummarizationSentenceCount >= 2 && 
                this.state.settings.textSummarizationSentenceCount <= 5 :
                true) &&
            !this.isNullOrWhiteSpace(this.state.settings.releaseMessageEndpoint);

        const contentsChanged: boolean = this.state.settings.azureTranslatorEnabled !== this.state.persistedSettings?.azureTranslatorEnabled ||
            this.state.settings.azureTranslatorKey !== this.state.persistedSettings?.azureTranslatorKey ||
            this.state.settings.azureTranslatorLocation !== this.state.persistedSettings?.azureTranslatorLocation ||
            this.state.settings.languageServiceEnabled !== this.state.persistedSettings?.languageServiceEnabled ||
            this.state.settings.languageServiceEndpoint !== this.state.persistedSettings?.languageServiceEndpoint ||
            this.state.settings.languageServiceKey !== this.state.persistedSettings?.languageServiceKey ||
            this.state.settings.textSummarizationSentenceCount !== this.state.persistedSettings?.textSummarizationSentenceCount ||
            this.state.settings.pirStorageUri !== this.state.persistedSettings?.pirStorageUri ||
            this.state.settings.releaseMessageEndpoint !== this.state.persistedSettings?.releaseMessageEndpoint ||
            this.state.settings.branding?.enabled !== this.state.persistedSettings?.branding?.enabled || 
            this.state.settings.branding?.logo !== this.state.persistedSettings?.branding?.logo ||
            this.state.settings.branding?.backgroundColor !== this.state.persistedSettings?.branding?.backgroundColor;

        return contentsValid && contentsChanged;
    }
}