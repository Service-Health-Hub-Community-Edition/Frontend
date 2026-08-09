import * as React from 'react';
import { ProgressIndicator } from '@fluentui/react';
import { Text, Button, Dialog, DialogType, DialogFooter } from '@fluentui/react';
import { GlobalState } from './GlobalState';
import { NotificationState } from './Notifications';
import { acquireAccessToken } from "../auth/AccessTokenHelper";
import { AuthenticationResult } from '@azure/msal-browser';

export interface ICustomAction {
    actionId: string;
    componentId: string;
    name: string;
    icon: string;
}

export interface ICustomActionState {
    state: CustomActionState,
    name: string,
    message: string
}

export enum CustomActionState {
    idle,
    running,
    success,
    error
}

export interface ICustomActionComponentState {
    componentName: string;
    customActions: ICustomAction[];
    customActionState: ICustomActionState;
}

export class CustomAction extends React.Component<{
    componentName: string,
    onLoad?: (actions: ICustomAction[]) => void
}, ICustomActionComponentState> {
    static contextType = GlobalState;

    constructor(props: {
        componentName: string,
        onLoad?: (actions: ICustomAction[]) => void
    }) {
        super(props);

        this.state = {
            componentName: props.componentName,
            customActions: [],
            customActionState: {
                state: CustomActionState.idle,
                name: "",
                message: ""
            }
        };
    }

    public render() {
        const {
            customActionState } = this.state;

        return (
            <div>
                <Dialog
                    hidden={customActionState.state === CustomActionState.idle}
                    onDismiss={this._closeRunCustomActionDialog}
                    dialogContentProps={{
                        type: DialogType.normal,
                        title: customActionState.name,
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

                    { customActionState.state === CustomActionState.running ? (<div className="customActionRunningProgress">
                        <ProgressIndicator label="Running" description={"Running custom action '" + customActionState.name + "'"} />
                    </div>) : (<Text variant={'medium'}>
                        <div dangerouslySetInnerHTML={{ __html: customActionState.message }} />
                    </Text>)}

                    <DialogFooter>
                        <Button onClick={this._closeRunCustomActionDialog} text="Close" />
                    </DialogFooter>
                </Dialog>
            </div>
        );
    }

    componentDidMount() {
        this._getCustomActionsFromApi(this.props.componentName);
    }

    componentDidUpdate(prevProps: { componentName?: string }) {
        if (prevProps.componentName !== this.props.componentName) {
            this._getCustomActionsFromApi(this.props.componentName);
        }
    }

    private _getCustomActionsFromApi(componentName?: string) {
        const requiredRoles: string[] = ['ServiceHealthReader', 'Communication.Write.All', 'Admin'];
        var authResponse: AuthenticationResult;
        var customActions: ICustomAction[] = [];

        if (componentName !== undefined && componentName.trim().length > 0) {
            acquireAccessToken()
                .then((response) => {
                    var tokenClaims: any = response.account?.idTokenClaims;
                    const userRoles: any = tokenClaims?.roles;
                    var userHasRequiredRole: boolean = userRoles.some((r: string) => requiredRoles.includes(r));
                    authResponse = response;

                    if (userHasRequiredRole)
                        fetch('/api/CustomAction?componentType=' + componentName, { headers: authResponse.idToken === "" ? {} : { 'Authorization': `Bearer ${authResponse.idToken}` } })
                            .then(response => {
                                if (response.ok) {
                                    return response.json();
                                }
                            })
                            .then(result => {
                                for (const customAction of result) {
                                    customActions.push({
                                        actionId: customAction.id,
                                        name: customAction.name,
                                        icon: customAction.icon,
                                        componentId: customAction.componentId
                                    });
                                }

                                customActions = _copyAndSort(customActions, 'name', false);

                                if (this.props.onLoad)
                                    this.props.onLoad(customActions);

                                this.setState({
                                    customActions: customActions
                                });
                            })
                }).catch((err) => {
                });
        }
    }

    _getCustomActions = (): ICustomAction[] => {
        if (this.state.customActions !== undefined && this.state.customActions !== null)
            return this.state.customActions;
        else
            return [];
    }
    _onRunCustomAction = (name: string, actionId: string, communicationId: string): void => {
        let globalState = this.context;
        const requiredRoles: string[] = ['ServiceHealthReader', 'Communication.Write.All', 'Admin'];
        var authResponse: AuthenticationResult;
        var params: RequestInit;
        var notificationId: string;

        acquireAccessToken()
            .then((response) => {
                var tokenClaims: any = response.account?.idTokenClaims;
                const userRoles: any = tokenClaims?.roles;
                var userHasRequiredRole: boolean = userRoles.some((r: string) => requiredRoles.includes(r));
                authResponse = response;

                if (userHasRequiredRole) {
                    const body = {
                        actionId: actionId,
                        communicationType: this.state.componentName,
                        communicationIds: [
                            communicationId
                        ]
                    };

                    params = {
                        headers: {
                            "Content-Type": "application/json charset=UTF-8",
                            "Authorization": "Bearer " + response.idToken
                        },
                        body: JSON.stringify(body),
                        method: "POST"
                    };

                    notificationId = globalState.setToastNotification({
                        id: '',
                        title: 'Running',
                        message: 'Started action "' + name + '" for item ' + communicationId + '.',
                        time: new Date(),
                        state: NotificationState.Running,
                        active: true
                    }, true);

                    this.setState({
                        customActionState: {
                            name: name,
                            state: CustomActionState.running,
                            message: "Running..."
                        }
                    });

                    var caResult: CustomActionState = CustomActionState.running;

                    fetch("/api/CustomAction/Run", params)
                        .then((response) => {

                            caResult = response.ok ? CustomActionState.success : CustomActionState.error;
                            return response.json();

                        })
                        .then(result => {
                            var message: string = "";

                            Object.keys(result).forEach(function (k) {
                                var color: string = result[k].isSuccessStatusCode ? "rgb(16, 124, 16)" : "#dc3545"
                                var icon: string = result[k].isSuccessStatusCode ? "CompletedSolid" : "AlertSolid"
                                message += "<b>" + k.toUpperCase() + "</b>&nbsp;<i class=\"ms-Icon ms-Icon--" + icon + "\" style=\"color: " + color + "\" aria-hidden=\"true\"></i></br>" + result[k].responseBody + "<br/><br/>";

                                globalState.setToastNotification({
                                    id: '',
                                    title: result[k].isSuccessStatusCode ? 'Completed' : 'Error',
                                    message: 'Action "' + name + '" ' + (result[k].isSuccessStatusCode ? 'completed' : 'failed') + ' for item ' + k.toUpperCase() + '. Result: ' + result[k].responseBody,
                                    time: new Date(),
                                    state: result[k].isSuccessStatusCode ? NotificationState.Success : NotificationState.Error,
                                    active: true
                                }, true);
                            });

                            if (caResult === CustomActionState.success)
                                globalState.removeNotification(notificationId)
                            else {
                                globalState.setToastNotification({
                                    id: notificationId,
                                    title: 'Error',
                                    message: 'Action "' + name + '" failed. Error: ' + message,
                                    time: new Date(),
                                    state: NotificationState.Error,
                                    active: true
                                }, true);
                            }

                            this.setState({
                                customActionState: {
                                    name: name,
                                    state: caResult,
                                    message: caResult === CustomActionState.success ? message : "Couldn't run action. Error" + message
                                }
                            });
                        });
                } else {
                    this.setState({
                        customActionState: {
                            name: name,
                            state: CustomActionState.error,
                            message: "Access denied. You don't have necessary permissions to run this action."
                        }
                    });
                }
            })
            .catch((err) => {
                this.setState({
                    customActionState: {
                        state: CustomActionState.error,
                        name: name,
                        message: "Authentication error. Details:<br/><br/>" + err.message
                    }
                });
            });;
    }  

    private _closeRunCustomActionDialog = (): void => {
        this.setState({
            customActionState: {
                state: CustomActionState.idle,
                name: "",
                message: ""
            }
        });
    }
}

function _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
    const key = columnKey as keyof T;
    return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
}