import * as React from 'react';
import {
    Selection, SelectionMode, IObjectWithKey, ICommandBarProps, IOverflowSetItemProps, Panel, PanelType, Dialog, DialogType, Link, ILinkStyleProps, ILinkStyles,
    DetailsList,
    CommandBar
} from '@fluentui/react';
import { DetailPageHeader } from '@m365-admin/detail-page';
import { IconAlert, IconAlertSize, IconAlertStatus } from '@m365-admin/icon-alert';
import { CompositeList, ICompositeListColumn, ICompositeListSelectionMap, ICompositeListSelectionMapItem, ICompositeListDetailsList } from '@m365-admin/composite-list';
import { EmptyStateImageSize, EmptyState } from '@m365-admin/empty-state';
import { M365Breadcrumb } from '@m365-admin/m365-breadcrumb';
import { AccessDenied } from "../../AccessDenied";
import { acquireAccessToken } from "../../../auth/AccessTokenHelper";
import { AuthenticationResult } from '@azure/msal-browser';
import { IConnector, IConnectorDefinition, IConnectorParameterDefinition, IConnectorConfigurationValue } from './Connectors';
import { RouteEditor } from './RouteEditor';

export interface IEntityProperty {
    displayName: string;
    entityProperty: string;
    type: string;
    hidden: boolean;
}

export interface IComponent {
    key?: string;
    id: string;
    name: string;
    internalName: string;
    icon?: string;
    capabilities: string[];
    entityProperties?: IEntityProperty[];
}

export interface IRoutingCondition {
    key?: string;
    order?: number;
    property: string;
    operator: string;
    value?: string;
    values?: string[];
    logicOperator?: string;
}

export interface IRoutingElement {
    key: string;
    id: string;
    order: number;
    name: string;
    conditions: IRoutingCondition[];
    icon: string;
    language?: string;
    stopProcessingOnMatch: boolean;
    hideWorkItemLink?: boolean;
    connector?: IConnector;
    connectorConfiguration: any;
}

interface IRoutingState {
    componentId: string;
    routeType: string;
    routing: IRoutingElement[];
    routingSelection: IObjectWithKey[];
    connectors: IConnector[];
    connectorDefinitions: IConnectorDefinition[];
    selectedRoute: IRoutingElement | undefined;
    component: IComponent | undefined;
    imageStore: IImage[];
    selectedId?: number;
    initialized: boolean;
    accessGranted: boolean;
    error?: string;
    saveError?: string;
    isEditRoutePanelOpen: boolean;
}

export interface IImage {
    id: number;
    name: string;
    type: string;
    format: string;
    content: string;
}

export class Routing extends React.Component<{ componentId: string, routeType: string }, IRoutingState> {  
    private selection = new Selection({
        onSelectionChanged: () => this._onRouteSelectionChanged()
    });

    constructor(props: {componentId: string, routeType: string}) {
        super(props);

        this.state = {
            componentId: this.props.componentId,
            routeType: this.props.routeType,
            routing: [],
            routingSelection: [],
            connectors: [],
            connectorDefinitions: [],
            selectedId: undefined,
            selectedRoute: undefined,
            component: undefined,
            imageStore: [],
            initialized: false,
            accessGranted: false,
            error: undefined,
            saveError: undefined,
            isEditRoutePanelOpen: false
        };
    }

    public render() {
        const {
            routing, component, initialized, routingSelection, selectedRoute, isEditRoutePanelOpen, connectors, connectorDefinitions
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
                key: 'ruleIcon',
                name: '',
                minWidth: 32,
                maxWidth: 32,
                fieldName: 'Icon',
                isResizable: false,
                isRowHeader: false,
                isIconOnly: true,
                onRender: (item: IRoutingElement) => {
                    return <div>
                        <img src={this._getImageData(item.icon)} style={{ width: '32px', marginTop: '-4px', float: 'left' }} />
                    </div>;
                }
            },
            {
                key: 'ruleName',
                name: 'Name',
                minWidth: 200,
                maxWidth: 350,
                fieldName: 'Name',
                isResizable: true,
                isRowHeader: true,
                calculatedWidth: 200,
                onRender: (item: IRoutingElement) => {
                    return <Link onClick={(event) => {
                        event.preventDefault();
                        this._onOpenEditRoutePanel(false);
                    }} styles={pipeFabricStyles}>{item.name}</Link >;
                }
            },
            {
                key: 'connector',
                name: 'Connector',
                minWidth: 75,
                maxWidth: 200,
                calculatedWidth: 100,
                isResizable: true,
                isCollapsable: true,
                isCollapsible: true, 
                
                onRender: (item: IRoutingElement) => {
                        return <div>
                            {item.connector?.icon ? (<img src={this._getImageData(item.connector?.icon)} style={{ width: '32px', marginTop: '-4px', float: 'left' }} />) : (<div/>)} {item.connector?.name}
                        </div>;
                }
            },
            {
                key: 'stopProcessingOnMatch',
                name: 'On match',
                minWidth: 120,
                maxWidth: 120,
                fieldName: 'StopProcessingOnMatch',
                isResizable: true,
                isCollapsable: true,
                isCollapsible: true,
                onRender: (item: IRoutingElement) => {
                    return <IconAlert
                        status={item.stopProcessingOnMatch ? IconAlertStatus.Blocked : IconAlertStatus.Success}
                        size={IconAlertSize.Small}
                        message={item.stopProcessingOnMatch ? ' Stop' : ' Continue'}
                    />
                }
            },
            {
                key: 'hideWorkItemLink',
                name: 'Show task link',
                minWidth: 120,
                maxWidth: 120,
                fieldName: 'HideWorkItemLink',
                isResizable: true,
                isCollapsable: true,
                isCollapsible: true,
                onRender: (item: IRoutingElement) => {
                    return <IconAlert
                        status={item.hideWorkItemLink ? IconAlertStatus.Error : IconAlertStatus.Success}
                        size={IconAlertSize.Small}
                    />
                }
            },
            {
                key: 'language',
                name: 'Language',
                minWidth: 120,
                maxWidth: 120,
                fieldName: 'language',
                isResizable: true,
                isCollapsable: true,
                isCollapsible: true
            }
        ];

        const commandBarProps: ICommandBarProps = {
            items: [
                {
                    key: 'newRoutingRule',
                    text: 'New',
                    iconProps: { iconName: 'Add' },
                    onClick: () => this._onOpenEditRoutePanel(true),
                },
                {
                    key: 'editRoutingRule',
                    text: 'Edit',
                    iconProps: { iconName: 'Edit' },
                    onClick: () => this._onOpenEditRoutePanel(false),
                    disabled: routingSelection !== undefined ? routingSelection.length <= 0 : true
                },
                {
                    key: 'deleteRule',
                    text: 'Delete',
                    iconProps: { iconName: 'Delete' },
                    onClick: () => window.confirm('This action will remove selected route. Proceed with caution.') ? this._onDeleteRoute() : this._doNothing(),
                    disabled: routingSelection !== undefined ? routingSelection.length <= 0 : true
                },
                {
                    key: 'moveUp',
                    text: 'Move up',
                    iconProps: { iconName: 'Up' },
                    onClick: () => this._move(true),
                    disabled: this._isMoveUpDownDisabled(true) 
                },
                {
                    key: 'moveDown',
                    text: 'Move down',
                    iconProps: { iconName: 'Down' },
                    onClick: () => this._move(false),
                    disabled: this._isMoveUpDownDisabled(false) 
                }
            ]
        };

        const contextMenuItems: IOverflowSetItemProps[] = [
            {
                key: 'moveUpCtx',
                text: 'Move up',
                title: 'Move up',
                'data-automation-id': 'moveUpCtx',
                iconProps: { iconName: 'Up' },
                onItemClick: (item: IRoutingElement): void => {
                    console.log('action - move up item', item);
                },
                onClick: () => {
                    console.log('Custom onClick that should still fire');
                }
            },
            {
                key: 'moveDownCtx',
                text: 'Move down',
                title: 'Move down',
                'data-automation-id': 'moveDownCtx',
                iconProps: { iconName: 'Down' },
                onItemClick: (item: IRoutingElement): void => {
                    console.log('action - move down item', item);
                },
                onClick: () => {
                    console.log('Custom onClick that should still fire');
                }
            },
            {
                key: 'deleteCtx',
                text: 'Delete',
                title: 'Delete',
                'data-automation-id': 'deleteCtx',
                iconProps: { iconName: 'Delete' },
                onItemClick: (item: IRoutingElement): void => {
                    console.log('action - delete item', item);
                },
                onClick: () => {
                    console.log('Custom onClick that should still fire');
                }
            },
        ];

        var listData: ICompositeListDetailsList[] = [
            {
                listProps: {
                    checkButtonAriaLabel: 'Select item',
                    ariaLabelForSelectAllCheckbox: 'Select all items',
                    items: routing,
                    columns: columns,
                    selectionMode: SelectionMode.single,
                    selection: this.selection,
                    setKey: 'id',
                    ariaLabelForGrid: 'Route list, use arrows to navigate'
          },
          key: 'routing'
        }

        ];

        return (
            <div className="container" style={{ maxWidth: '97%' }}>
                <div className="row">
                    <div className="col">
                        <M365Breadcrumb
                            items={[
                                { text: 'Home', key: 'home', href: '/admin' },
                                { text: 'Notifications', key: 'notifications', href: '/admin/notifications' },
                                { text: 'Components', key: 'components', href: '/admin/notifications/entities' },
                                { text: component ? component.name : 'Undefined', key: component ? component.internalName : 'undefined' },
                                { text: 'Routing', key: 'routing', isCurrentItem: true }
                            ]}
                            style={{ marginBottom: '16px' }}
                        />
                        <DetailPageHeader
                            title="Routing"
                            description="Notifications will be distributed throughout different connectors based on the routing rules and order defined in the table below. A catch-all rule should be also created and placed at the end of the routing table. New rules will be created right before the last rule."
                        />
                    </div>
                </div>
                <div className="row">
                    <div className="col">
                        <CommandBar
                            items={commandBarProps.items}
                        />
                            <DetailsList
                                checkButtonAriaLabel='Select item'
                                ariaLabelForSelectAllCheckbox='Select all items'
                                items={routing}
                                columns={columns}
                                selectionMode={SelectionMode.single}
                                selection={this.selection}
                                setKey='id'
                                ariaLabelForGrid='Route list, use arrows to navigate'
                        />

                        <Panel
                            headerText={selectedRoute ? selectedRoute.name : "New route"}
                            isOpen={isEditRoutePanelOpen}
                            onDismiss={this._onDismisEditRoutePanel}
                            type={PanelType.medium}
                            isBlocking={true}
                            allowTouchBodyScroll={true}
                            // minWidth={350}
                            // You MUST provide this prop! Otherwise screen readers will just say "button" with no label.
                            closeButtonAriaLabel="Close"
                        /* styles={{
                            main: {
                                selectors: {
                                    ['@media (min-width: 480px)']: {
                                        width: 1440,
                                        minWidth: 350,
                                        maxWidth: '2000px'
                                    }
                                }
                            }
                        }} */
                        >
                            <RouteEditor
                                route={selectedRoute}
                                component={component}
                                connectors={connectors}
                                connectorDefinitions={connectorDefinitions}
                                onSave={this._onCommitRoute}
                                onCancel={this._onDismisEditRoutePanel}
                            />
                        </Panel>
                        
                        </div>
                </div>
                <div className="row">
                    <div className="col">
                        {routing === undefined ||
                            routing.length === 0 ? (
                            <EmptyState
                                    title='No rules found'
                                    body='Create at least one routing rule to enable notifications.'
                                    imageProps={{
                                        src: '/images/empty.png',
                                        alt: 'No routing rules graphic'
                                    }}
                                    illustrationSize={EmptyStateImageSize.Small}

                                    actionBarProps={{
                                        primaryButtonProps: {
                                            text: 'New rule',
                                            onClick: () => this._onOpenEditRoutePanel(true)
                                        }
                                    }}
                                    styles={{ root: { justifyContent:'' } }}
                                />) : ""}
                    </div>
                </div>
            </div>
        );
    }

    componentDidMount() {
        this._onLoadConfig();
    }

    private _onLoadConfig(): void {
        const requiredRoles: string[] = ['Admin'];
        var authResponse: AuthenticationResult;
        var userHasRequiredRole: boolean = false;
        var routeCollection: IRoutingElement[] = [];
        var connectorDefinitions: IConnectorDefinition[] = [];
        var connectors: IConnector[] = [];
        var connectorDefinitions: IConnectorDefinition[] = [];
        var component: IComponent | undefined = undefined;
        var imageStore: IImage[] = [];

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
            }).then(() => {
                if (userHasRequiredRole)
                    fetch('/api/ConnectorDefinition?type='+this.state.routeType, { headers: authResponse.idToken === "" ? {} : { 'Authorization': `Bearer ${authResponse.idToken}` } })
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
                            for (const connectorDefinition of result) {
                                if (!connectorDefinition.hidden)
                                    connectorDefinitions.push({
                                        name: connectorDefinition.name,
                                        id: connectorDefinition.connectorId,
                                        parameters: connectorDefinition.parameters,
                                        type: connectorDefinition.type,
                                        icon: connectorDefinition.icon,
                                        system: connectorDefinition.system,
                                        unique: connectorDefinition.unique
                                    });
                            }

                            this.setState({
                                connectorDefinitions: connectorDefinitions
                            });
                        }).catch((err) => {
                            this.setState({
                                error: err.message
                            });
                        });
            })
            .then(() => {
                if (userHasRequiredRole)
                    fetch('/api/ImageStore?type=notificationIcon', { headers: authResponse.idToken === "" ? {} : { 'Authorization': `Bearer ${authResponse.idToken}` } })
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
                            for (const image of result) {
                                imageStore.push({
                                    id: image.id,
                                    name: image.name,
                                    type: image.type,
                                    format: image.format,
                                    content: image.content
                                });
                            }

                            this.setState({
                                imageStore: imageStore
                            });
                        }).catch((err) => {
                            this.setState({
                                error: err.message
                            });
                        });
            })
            .then(() => {
                if (userHasRequiredRole)
                    fetch('/api/Connectors?type=notificationManager', { headers: authResponse.idToken === "" ? {} : { 'Authorization': `Bearer ${authResponse.idToken}` } })
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
                            for (const connector of result) {
                                if (!connector.hidden)
                                    connectors.push({
                                        key: connector.connectorId,
                                        name: connector.name,
                                        connectorId: connector.connectorId,
                                        configuration: connector.configuration,
                                        parameterDefinition: connector.parameterDefinition,
                                        type: connector.type,
                                        icon: connector.icon,
                                        connectorType: connector.connectorType,
                                        connectorTypeName: connector.connectorTypeName,
                                        system: connector.system
                                    });
                            }

                            connectors = _copyAndSort(connectors, 'name', false);
                            console.log(connectors);
                            this.setState({
                                connectors: connectors
                            });
                        }).catch((err) => {
                            this.setState({
                                error: err.message
                            });
                        });
            })
            .then(() => {
                if (userHasRequiredRole)
                    fetch('/api/Components?id='+this.state.componentId, { headers: authResponse.idToken === "" ? {} : { 'Authorization': `Bearer ${authResponse.idToken}` } })
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
                            if (result !== undefined && result.length > 0)
                                component = result[0] as IComponent;

                            this.setState({
                                component: component
                            });
                        }).catch((err) => {
                            this.setState({
                                error: err.message
                            });
                        });
            })
            .then(() => {
                if (userHasRequiredRole)
                    fetch('/api/Route?type=' + this.state.routeType + '&component='+this.state.componentId, { headers: authResponse.idToken === "" ? {} : { 'Authorization': `Bearer ${authResponse.idToken}` } })
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
                                routeCollection.push({
                                    key: route.id,
                                    id: route.id,
                                    name: route.name,
                                    order: route.order,
                                    icon: route.icon,
                                    conditions: route.conditions,
                                    stopProcessingOnMatch: route.stopProcessingOnMatch,
                                    hideWorkItemLink: route.hideWorkItemLink,
                                    language: route.language,
                                    connector: route.connector,
                                    connectorConfiguration: route.connectorConfiguration
                                });
                            }

                            routeCollection = _copyAndSort(routeCollection, 'order', false);

                            this.setState({
                                routing: routeCollection,
                                initialized: true
                            });
                        }).catch((err) => {
                            this.setState({
                                error: err.message
                            });
                        });
            });
    }

    private _getImageData(url: string) {
        if (url.toLowerCase().startsWith('imagestore://')) {
            var name: string = url.slice(13); // remove imagestore:// from start of the string
            var image: IImage | undefined = this.state.imageStore.find((i) => i.name.toLowerCase() === name.toLowerCase());
            if (image !== undefined) {
                return 'data:' + image.format + ';base64,' + image.content
            } else
                return '';
        } else
            return url;
    }

    private _createGuid(): string {
        return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
            var r = Math.random() * 16 | 0, v = c === 'x' ? r : (r & 0x3 | 0x8);
            return v.toString(16);
        });
    }

    private _onCommitRoute = (route: IRoutingElement): void => {
        var routing: IRoutingElement[] = this.state.routing;
        var oldRoute: IRoutingElement | undefined = routing.find((r) => r.key === route.key);
        var newRoute: IRoutingElement;

        if (oldRoute !== undefined) {
            oldRoute.conditions = route.conditions;
            oldRoute.connector = route.connector;
            oldRoute.connectorConfiguration = route.connectorConfiguration;
            oldRoute.hideWorkItemLink = route.hideWorkItemLink;
            oldRoute.icon = route.icon;
            oldRoute.id = route.id;
            oldRoute.language = route.language;
            oldRoute.name = route.name;
            oldRoute.stopProcessingOnMatch = route.stopProcessingOnMatch;

            this._onSaveRoute(route);
        } else {
            
            var nextOrderId: number = 0;
            var orderedRoutes = _copyAndSort(routing, 'order', true);
            if (orderedRoutes !== undefined && orderedRoutes.length > 0) {
                var lastRoute = orderedRoutes[0];
                nextOrderId = lastRoute.order ? lastRoute.order : 0;

                // move last route one place lower to make space for a new route. New route will be created before catch all rule
                routing = routing.filter((r) => r.id !== lastRoute.id);
                lastRoute.order += 1;
                routing.push(lastRoute);
                this._onSaveRoute(lastRoute);
            }

            var newId: string = this._createGuid();
            newRoute = {
                id: newId,
                key: newId,
                name: route.name,
                language: route.language,
                order: nextOrderId,
                conditions: route.conditions,
                connector: route.connector,
                connectorConfiguration: route.connectorConfiguration,
                icon: route.icon,
                stopProcessingOnMatch: route.stopProcessingOnMatch,
                hideWorkItemLink: route.hideWorkItemLink
            };

            routing.push(newRoute);
            this._onSaveRoute(newRoute);
        }

        routing = _copyAndSort(routing, 'order', false);
        this.setState({
            routing: routing,
            isEditRoutePanelOpen: false
        });
    }

    private _onSaveRoute = (route: IRoutingElement): void => {
        var params: RequestInit;

        var routeBody = {
            id: route.id,
            name: route.name,
            order: route.order,
            icon: route.icon,
            conditions: route.conditions,
            stopProcessingOnMatch: route.stopProcessingOnMatch,
            hideWorkItemLink: route.hideWorkItemLink,
            language: route.language,
            connector: route.connector?.connectorId,
            connectorConfiguration: route.connectorConfiguration,
            component: this.state.componentId
        };

        acquireAccessToken()
            .then((response) => {
                params = {
                    headers: {
                        "Content-Type": "application/json charset=UTF-8",
                        "Authorization": "Bearer " + response.idToken
                    },
                    body: JSON.stringify(routeBody),
                    method: "POST"
                };

                fetch("/api/Route", params)
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

    private _onDeleteRoute = (): void => {
        var routing = this.state.routing;
        var route = routing.find((r: IRoutingElement) => r.id === this.state.selectedRoute?.id);
        var params: RequestInit;

        if (route !== undefined) {
            acquireAccessToken()
                .then((response) => {

                    const routeBody: any = {
                        id: route!.id
                    }

                    params = {
                        headers: {
                            "Content-Type": "application/json charset=UTF-8",
                            "Authorization": "Bearer " + response.idToken
                        },
                        body: JSON.stringify(routeBody),
                        method: "DELETE"
                    };

                    fetch("/api/Route", params)
                        .then((response) => {
                            if (routing.find((r: IRoutingElement) => r.id === route!.id) !== undefined)
                                routing = routing.filter((r: IRoutingElement) => r.id !== route!.id);

                            this.setState({
                                routing: routing,
                                routingSelection: [],
                                selectedRoute: undefined,
                                selectedId: undefined,
                                isEditRoutePanelOpen: false
                            });
                        })
                        .catch((err) => {
                            this.setState({
                                isEditRoutePanelOpen: false,
                                saveError: err.message
                            });
                        });
                })
                .catch((err) => {
                    this.setState({
                        isEditRoutePanelOpen: false,
                        saveError: "Authentication error. Details:<br/><br/>" + err.message
                    });
                });
        }
    }

    private _doNothing = (): void => {
        
    }

    _onOpenEditRoutePanel(newItemMode: boolean) {
        if (newItemMode) 
            this.setState({
                selectedRoute: undefined,
                routingSelection: [],
                isEditRoutePanelOpen: true
            });
        else
            this.setState({
                isEditRoutePanelOpen: true
            });
    }

    private _onDismisEditRoutePanel = (): void => {
        this.setState({
            isEditRoutePanelOpen: false
        });
    }

    private _onSelectionChanged = (selectionMap: ICompositeListSelectionMap, listKey: string) => {
        const activeSelection: ICompositeListSelectionMapItem | undefined = selectionMap.get(listKey);
        const newSelection: IObjectWithKey[] | undefined = activeSelection?.selection.getSelection();
        var selectedRoute: IRoutingElement | undefined = newSelection === undefined ? undefined : newSelection[0] as IRoutingElement;
       
        if (newSelection) {
            this.setState({
                selectedRoute: selectedRoute,
                routingSelection: newSelection
            });
        } else {
            this.setState({
                selectedRoute: undefined,
                routingSelection: []
            });
        }
    };

    private _onRouteSelectionChanged = () => {
        const items: IObjectWithKey[] = this.selection.getSelection();
        var selectedRoute: IRoutingElement | undefined = items === undefined || items.length <=0 ? undefined : items[0] as IRoutingElement;

        if (selectedRoute) {
            this.setState({
                selectedRoute: selectedRoute,
                routingSelection: items
            });
        } else {
            this.setState({
                selectedRoute: undefined,
                routingSelection: []
            });
        }
    };

    private _isMoveUpDownDisabled = (first: boolean): boolean => {
        if (this.state.routingSelection !== undefined && this.state.routingSelection.length > 0 &&
            this.state.routing !== undefined && this.state.routing.length > 0) {
            var orderedRules = _copyAndSort(this.state.routing, 'order', !first);
            var firstRule = orderedRules[0];
            var selection: IRoutingElement = this.state.routingSelection[0] as IRoutingElement;
            return selection ? selection.order === firstRule.order : false;
        } else {
            return true;
        }
    }

    private _move = (up: boolean): void => {
        if (this.state.routingSelection !== undefined && this.state.routingSelection.length > 0) {
            var selection: IRoutingElement = this.state.routingSelection[0] as IRoutingElement;
            var selectedOrderValue: number = selection.order;
            var currentRule = this.state.routing.find(element => element.order === selectedOrderValue);
            if (currentRule !== undefined) {
                var index = this.state.routing.indexOf(currentRule!);
                var movingRule: IRoutingElement | undefined = undefined;
                if (up) {
                    if (index > 0)
                        movingRule = this.state.routing[index - 1];
                } else {
                    if (index < this.state.routing.length-1)
                        movingRule = this.state.routing[index + 1];
                }

                if (movingRule !== undefined) {
                    currentRule.order = movingRule.order;
                    movingRule.order = selectedOrderValue;
                    var routing = this.state.routing;

                    this._onSaveRoute(currentRule);
                    this._onSaveRoute(movingRule);

                    routing.sort((a: IRoutingElement, b: IRoutingElement) => a.order > b.order ? 1 : -1);

                    this.selection.setItems(routing);

                    this.state.routingSelection.map(r => {
                        if (r.key !== undefined && r.key !== null)
                            this.selection.setKeySelected(r.key as string, true, false)
                    });

                    this.setState({
                        routing: routing
                    });
                }
            }
        }
    }
}

function _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
    const key = columnKey as keyof T;
    return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
}