import * as React from 'react';
import {
    Text, Spinner, SpinnerSize, MessageBar, MessageBarType,
    CommandBar, ICommandBarItemProps, Dialog, DialogType, DialogFooter,
    PrimaryButton
} from '@fluentui/react';

import { FontIcon, mergeStyles, mergeStyleSets } from '@fluentui/react';
import Editor from "@monaco-editor/react";
import { AccessDenied } from "../../AccessDenied";
import { acquireAccessToken } from "../../../auth/AccessTokenHelper";
import { AuthenticationResult } from '@azure/msal-browser';

interface IConfigEditorState {
    component: string;
    element: string;
    value: string;
    newValue: string;
    isDataLoaded: boolean;
    error?: string;
    saveError?: string;
    saveSuccessDialogOpen: boolean;
    accessGranted?: boolean;
    editor?: any;
}

export class ConfigEditor extends React.Component<{ component: string, element: string }, IConfigEditorState> {
    constructor(props: { component: string, element: string }) {
        super(props);

        this.state = {
            component: this.props.component,
            element: this.props.element,
            value: "",
            newValue: "",
            isDataLoaded: false,
            error: undefined,
            saveError: undefined,
            saveSuccessDialogOpen: false,
            accessGranted: undefined,
            editor: undefined
        };

    }

    public render() {
        const {
            value, newValue, isDataLoaded, error, saveError, saveSuccessDialogOpen, accessGranted
        } = this.state;

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

        if (!isDataLoaded) {
            return (
                <div className="container" style={{ maxWidth: '1400px' }}>
                    <div className="row">
                        <div className="col">
                            <br />&nbsp;
                        </div>
                    </div>
                    <div className="row">
                        <div className="col">
                            <div className="loadingProgress" style={{ display: isDataLoaded || error !== undefined ? 'none' : 'block' }}>
                                <br /><br />
                                <Spinner size={SpinnerSize.large} />
                            </div>
                            <div style={{ display: error != undefined ? 'block' : 'none' }}>
                                <MessageBar
                                    messageBarType={MessageBarType.error}
                                    isMultiline={false}
                                >
                                    {"Couldn't retrieve data. Error: " + error}
                                </MessageBar>
                                <br />
                            </div>
                        </div>
                    </div>
                </div>

            );
        }

        const _commandBarItems: ICommandBarItemProps[] = [
            {
                key: 'save',
                text: 'Save',
                cacheKey: 'cacheSave', // changing this key will invalidate this item's cache
                iconProps: { iconName: 'Save' },
                onClick: () => this._onSaveConfig(),
                disabled: value === newValue
            },
            {
                key: 'reload',
                text: 'Reload',
                cacheKey: 'cacheReload', // changing this key will invalidate this item's cache
                iconProps: { iconName: 'Refresh' },
                onClick: () => this._onLoadConfig()
            }
        ];

        const _commandBarFarItems: ICommandBarItemProps[] = [];

        return (
            <div style={{ border: "1px solid rgb(225, 225, 225)" }}>
                <div>
                    <CommandBar items={_commandBarItems} farItems={_commandBarFarItems} />
                </div>
                <div className="configEditor">
                    <Editor
                        height="58vh"
                        language="json"
                        value={value}
                        options={{ automaticLayout: true }}
                        onChange={this._handleEditorChange}
                        onMount={this._handleEditorOnMount}
                    />
                </div>

                <Dialog
                    hidden={saveError === undefined}
                    onDismiss={this._closeSaveErrorDialog}
                    dialogContentProps={{
                        type: DialogType.normal,
                        title: 'Error',
                    }}
                    modalProps={{
                        isBlocking: true
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
                    <Text variant={'medium'}>
                        <div dangerouslySetInnerHTML={{ __html: saveError! }} />
                    </Text>

                    <DialogFooter>
                        <PrimaryButton onClick={this._closeSaveErrorDialog} text="Close" />
                    </DialogFooter>
                </Dialog>

                <Dialog
                    hidden={!saveSuccessDialogOpen}
                    onDismiss={this._closeSaveSuccessDialog}
                    dialogContentProps={{
                        type: DialogType.normal,
                        title: 'Configuration saved',
                    }}
                    modalProps={{
                        isBlocking: true
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
                    <Text variant={'medium'}>
                        Configuration successfully saved.
                    </Text>

                    <DialogFooter>
                        <PrimaryButton onClick={this._closeSaveSuccessDialog} text="Close" />
                    </DialogFooter>
                </Dialog>
            </div>
        );
    }

    componentDidMount(): void {
        window.addEventListener('resize', this._handleResize);
        this._onLoadConfig();
    }

    private _handleEditorChange = (value: any, event: any): void => {
        this.setState({
            newValue: value !== undefined ? value : ""
        });
    }

    private _handleEditorOnMount = (editor: any, monaco: any): void => {
        this.setState({
            editor: editor
        });
    }

    private _handleResize = () => {
        if (this.state.editor !== undefined) {
            this.state.editor.layout();
        }
    }

    private _onLoadConfig(): void {
        const requiredRoles: string[] = ['Admin'];
        var authResponse: AuthenticationResult;
        var userHasRequiredRole: boolean = false;

        this.setState({
            isDataLoaded: false
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
            }).then(() => {
                if (userHasRequiredRole)
                    fetch('/api/SyncConfig?component=' + this.state.component + '&element=' + this.state.element, { headers: authResponse.idToken === "" ? {} : { 'Authorization': `Bearer ${authResponse.idToken}` } })
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
                        .then(result => {
                            this.setState({
                                component: result.component,
                                element: result.element,
                                value: result.config,
                                newValue: result.config,
                                isDataLoaded: true
                            });
                        }).catch((err) => {
                            this.setState({
                                error: err.message
                            });
                        });
            });
    }

    private _onSaveConfig(): void {
        var params: RequestInit;
        acquireAccessToken()
            .then((response) => {
                const body = {
                    component: this.state.component,
                    element: this.state.element,
                    config: this.state.newValue
                };

                params = {
                    headers: {
                        "Content-Type": "application/json charset=UTF-8",
                        "Authorization": "Bearer " + response.idToken
                    },
                    body: JSON.stringify(body),
                    method: "POST"
                };

                fetch("/api/SyncConfig", params)
                    .then((response) => {
                        if (!response.ok) {
                            // make the promise be rejected if we didn't get a 2xx response
                            const err = new Error("Couldn't update configuration. Error details:<br/><br/><b>HTTP " + response.status + "</b><br/>" + response.statusText);
                            throw err;
                        }

                        this.setState({
                            saveSuccessDialogOpen: true
                        });
                    })
                    .catch((err) => {
                        this.setState({
                            saveError: err.message
                        });
                    });
            })
            .catch((err) => {
                this.setState({
                    saveError: "Authentication error. Details:<br/><br/>" + err.message
                });
            });
    }

    private _closeSaveErrorDialog = (): void => {
        this.setState({
            saveError: undefined
        });
    }

    private _closeSaveSuccessDialog = (): void => {
        this.setState({
            saveSuccessDialogOpen: false
        });
    }
}