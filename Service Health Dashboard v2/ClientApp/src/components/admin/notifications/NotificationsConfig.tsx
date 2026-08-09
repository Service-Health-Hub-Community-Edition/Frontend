import * as React from 'react';
import {
    SelectionMode, IObjectWithKey, ICommandBarProps,
    Panel, PanelType, Dialog, DialogType, TextField, Link, ILinkStyleProps, ILinkStyles,
    PrimaryButton, DefaultButton, DialogFooter, Label, Dropdown, IDropdownOption, ISelectableOption, Text, TextStyles
} from '@fluentui/react';
import { DetailPageHeader } from '@m365-admin/detail-page';
import { IconAlert, IconAlertSize, IconAlertStatus } from '@m365-admin/icon-alert';
import { CompositeList, ICompositeListColumn, ICompositeListSelectionMap, ICompositeListSelectionMapItem } from '@m365-admin/composite-list';
import { EmptyStateImageSize } from '@m365-admin/empty-state';
import { M365Breadcrumb } from '@m365-admin/m365-breadcrumb';
import { AccessDenied } from "../../AccessDenied";
import { acquireAccessToken } from "../../../auth/AccessTokenHelper";
import { AuthenticationResult } from '@azure/msal-browser';
import { IComponent, IRoutingElement } from './Routing';
import { AdminLayout } from '../AdminLayout';

interface IConfigAction {
    key: string;
    name: string;
    status: string;
    description: string;
    link: string;
}

interface INotificationsConfigState {
    actions: IConfigAction[];
    routingConfigState: string;
    initialized: boolean;
    accessGranted: boolean;
    error?: string;
}

export class NotificationsConfig extends React.Component<{}, INotificationsConfigState> {  

    constructor(props: {}) {
        super(props);

        this.state = {
            actions: [],
            routingConfigState: '',
            initialized: false,
            accessGranted: false,
            error: undefined
        };
    }

    public render() {
        const {
            routingConfigState, initialized
        } = this.state;

        if (!initialized)
            return (<div/>);

        const pipeFabricStyles = (p: ILinkStyleProps): ILinkStyles => ({
            root: {
                textDecoration: 'none',
                color: p.theme.semanticColors.bodyText
            },
        });

        const columns: ICompositeListColumn[] = [
            {
                key: 'actionName',
                name: 'Name',
                minWidth: 150,
                fieldName: 'name',
                isResizable: true,
                isRowHeader: true,
                onRender: (item: IConfigAction) => {
                    return <Link href={item.link} styles={pipeFabricStyles}>{item.name}</Link >;
                }
            },
            {
                key: 'status',
                name: 'Status',
                minWidth: 120,
                fieldName: 'status',
                isResizable: true,
                onRender: (item: IConfigAction) => {
                    switch (item.status) {
                        case 'Completed':
                            return <div style={{ display: 'inline-block' }} ><IconAlert
                                status={IconAlertStatus.Success}
                                size={IconAlertSize.Small}
                                message={' ' + item.status}                               
                            />
                                
                            </div>

                        case '':
                            return <div/>

                        case 'Not started yet':
                            return <IconAlert
                                status={IconAlertStatus.NotStarted}
                                size={IconAlertSize.Small}
                                message={item.status}
                            />

                        default:
                            return <IconAlert
                                status={IconAlertStatus.Warning}
                                size={IconAlertSize.Small}
                                message={item.status}
                            />
                    }                 
                }
            },
            {
                key: 'description',
                name: 'Description',
                minWidth: 120,
                maxWidth: 500,
                isMultiline: true,
                fieldName: 'description',
                isResizable: true
            }
        ];

        const commandBarProps: ICommandBarProps = {
            items: []
        };

        var actions: IConfigAction[] = [
            {
                key: 'actionConnections',
                name: 'Get your connections set up',
                status: '',
                description: 'Connecting logic apps will allow us to send the notification to different systems.',
                link: '/admin/notifications/connectors'
            },
            {
                key: 'actionRouting',
                name: 'Configure routing',
                status: routingConfigState,
                description: 'Defining notification routing rules will help target relevant users, informing them about the events within Microsoft tenant.',
                link: '/admin/notifications/entities'
            }
        ];

        var listData = [
            {
                listProps: {
                    checkButtonAriaLabel: 'Select item',
                    ariaLabelForSelectAllCheckbox: 'Select all items',
                    items: actions,
                    columns: columns,
                    selectionMode: SelectionMode.none,
                    ariaLabelForGrid: 'Actions list, use arrows to navigate'
                },
                key: 'actions'
            }
        ];
        
        return (
            <AdminLayout>
            <div className="container" style={{ maxWidth: '97%' }}>
                <div className="row">
                    <div className="col">
                        <M365Breadcrumb
                            items={[
                                { text: 'Home', key: 'home', href: '/admin' },
                                { text: 'Notifications', key: 'notifications', isCurrentItem: true }
                            ]}
                            style={{ marginBottom: '16px' }}
                        />
                        <DetailPageHeader
                            title="Notification settings"
                            description="Configure notifications by defining connections and routing for Microsoft Service Health Hub components."
                        />
                    </div>
                </div>
                <div className="row">
                    <div className="col">
                        <CompositeList
                            detailsListProps={listData}
                            commandBarProps={commandBarProps}
                            defaultIsCompactMode={true}
                            onSelectionChanged={this._doNothing}
                            legacyInvokeAndSelect={true}
                            isEmpty={
                                listData === undefined ||
                                listData[0].listProps?.items === undefined ||
                                listData[0].listProps?.items?.length === 0 }
                            emptyStateProps={{
                                title: 'No actions found',
                                body: 'Please contact your administrator.',
                                imageProps: {
                                    src: '/images/empty.png',
                                    alt: 'No actions found graphic'
                                },
                                illustrationSize: EmptyStateImageSize.Large
                            }}
                        />
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
                if (userHasRequiredRole) {
                    this._getRoutingState(authResponse);
                    this.setState({
                        initialized: true
                    });
                }
            }).catch((err) => {
                this.setState({
                    error: err.message
                });
            });
    }

    private _doNothing = (): void => {
        
    }

    private _getRoutingState(authResponse: AuthenticationResult) {
        interface IComponentState {
            id: string;
            configPresent: boolean;
        }

        var components: IComponentState[] = [];
        var currentState: string = '';

        fetch('/api/Components', { headers: authResponse.idToken === "" ? {} : { 'Authorization': `Bearer ${authResponse.idToken}` } })
            .then(response => {
                if (response.ok) {
                    return response.json();
                } else {
                    this.setState({
                        error: response.status + " " + response.statusText
                    });
                    throw Error(response.status + " " + response.statusText);
                }
            }).then(result => {
                for (const c of result)
                    if (c.capabilities?.includes('Routing') && c.capabilities?.includes('Notifications'))
                        components.push({
                            id: c.id,
                            configPresent: false
                        });
            }).then(() => {
                    fetch('/api/Route?type=notificationManager', { headers: authResponse.idToken === "" ? {} : { 'Authorization': `Bearer ${authResponse.idToken}` } })
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
                            for (const route of result) {
                                var component = components.find(c => c.id === route.component);
                                if (component !== undefined)
                                    component.configPresent = true;
                            }

                            var missingConfig = components.filter(c => c.configPresent === false);

                            if (missingConfig !== undefined && missingConfig.length > 0)
                                currentState = 'Missing configuration'
                            else
                                currentState = 'Completed'

                            this.setState({
                                routingConfigState: currentState
                            });
                            
                        }).catch((err) => {
                            this.setState({
                                error: err.message
                            });
                        });
            })
            .catch((err) => {
                this.setState({
                    error: err.message
                });
            });
    }
}

function _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
    const key = columnKey as keyof T;
    return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
}