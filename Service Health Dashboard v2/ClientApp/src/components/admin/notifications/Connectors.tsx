import * as React from 'react';
import {
    SelectionMode, IObjectWithKey, ICommandBarProps,
    Panel, PanelType, Dialog, DialogType, TextField, Link, ILinkStyleProps, ILinkStyles,
    PrimaryButton, DefaultButton, DialogFooter, Label, Dropdown, IDropdownOption, ISelectableOption
} from '@fluentui/react';
import { DetailPageHeader } from '@m365-admin/detail-page';
import { IconAlert, IconAlertSize, IconAlertStatus } from '@m365-admin/icon-alert';
import { CompositeList, ICompositeListColumn, ICompositeListSelectionMap, ICompositeListSelectionMapItem } from '@m365-admin/composite-list';
import { EmptyStateImageSize } from '@m365-admin/empty-state';
import { M365Breadcrumb } from '@m365-admin/m365-breadcrumb';
import { AccessDenied } from "../../AccessDenied";
import { AdminLayout } from '../AdminLayout';
import { acquireAccessToken } from "../../../auth/AccessTokenHelper";
import { AuthenticationResult } from '@azure/msal-browser';

export interface IConnectorConfigurationValue {
    name: string;
    value: any;
}

export interface IConnector {
    key: string;
    connectorId: string;
    name: string;
    connectorType: string;
    connectorTypeName: string;
    type: string;
    icon: string;
    configuration: IConnectorConfigurationValue[];
    parameterDefinition: IConnectorParameterDefinition[];
    system: boolean;
}

export interface IConnectorParameterDefinition {
    id: number;
    name: string;
    displayName: string;
    type: string;
    description: string;
    parameterType: string;
}

export interface IConnectorDefinition {
    id: string;
    name: string;
    type: string;
    icon: string;
    system: boolean;
    unique: boolean;
    parameters: IConnectorParameterDefinition[];
}

interface IConnectorsState {
    connectors: IConnector[];
    connectorSelection: IObjectWithKey[];
    connectorEditorValues: IConnectorConfigurationValue[];
    selectedConnector: IConnector | undefined;
    selectedConnectorParameterDefinition: IConnectorParameterDefinition[];
    selectedConnectorConfiguration: IConnectorConfigurationValue[];
    connectorDefinitions: IConnectorDefinition[];
    isEditConnectorPanelOpen: boolean;
    editorPanelNewItemMode: boolean;
    newItemConnectorType?: string;
    newConnectorName?: string;
    selectedId?: string;
    initialized: boolean;
    accessGranted: boolean;
    error?: string;
    saveError?: string;
}

export class Connectors extends React.Component<{}, IConnectorsState> {  

    constructor(props: {}) {
        super(props);

        this.state = {
            connectors: [],
            connectorSelection: [],
            connectorEditorValues: [],
            selectedConnector: undefined,
            selectedConnectorParameterDefinition: [],
            selectedConnectorConfiguration: [],
            connectorDefinitions: [],
            isEditConnectorPanelOpen: false,
            editorPanelNewItemMode: false,
            newItemConnectorType: undefined,
            newConnectorName: undefined,
            selectedId: undefined,
            initialized: false,
            accessGranted: false,
            error: undefined,
            saveError: undefined
        };
    }

    public render() {
        const {
            connectors, initialized, connectorSelection,
            selectedConnector, selectedConnectorParameterDefinition,
            connectorEditorValues, isEditConnectorPanelOpen, connectorDefinitions,
            editorPanelNewItemMode, newConnectorName, newItemConnectorType
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
                key: 'connectorIcon',
                name: '',
                minWidth: 32,
                maxWidth: 32,
                fieldName: 'icon',
                isResizable: false,
                isRowHeader: true,
                isIconOnly: true,
                onRender: (item: IConnector) => {
                    return <div>
                        <img src={item.icon} style={{ width: '32px', marginTop: '-4px', float: 'left' }} />
                    </div>;
                }
            },
            {
                key: 'connectorName',
                name: 'Name',
                minWidth: 150,
                maxWidth: 350,
                fieldName: 'name',
                isResizable: true,
                isRowHeader: true,
                onRender: (item: IConnector) => {
                    return <Link onClick={(event) => {
                        event.preventDefault();
                        this._onOpenEditConnectorPanel(false);
                    }} styles={pipeFabricStyles}>{item.name}</Link >;
                }
            },
            {
                key: 'connectorId',
                name: 'Id',
                minWidth: 280,
                maxWidth: 280,
                fieldName: 'id',
                isResizable: false,
                isRowHeader: true,
                onRender: (item: IConnector) => {
                    return <div>
                        {item.connectorId}
                    </div>;
                }
            },
            {
                key: 'connectorType',
                name: 'Type',
                minWidth: 200,
                maxWidth: 200,
                fieldName: 'type',
                isResizable: false,
                isRowHeader: true,
                onRender: (item: IConnector) => {
                    return <div>
                        {item.connectorTypeName}
                    </div>;
                }
            }
        ];

        const commandBarProps: ICommandBarProps = {
            items: [
                {
                    key: 'newConnector',
                    text: 'New',
                    iconProps: { iconName: 'Add' },
                    onClick: () => this._onOpenEditConnectorPanel(true),
                },
                {
                    key: 'editConnector',
                    text: 'Edit',
                    iconProps: { iconName: 'Edit' },
                    onClick: () => this._onOpenEditConnectorPanel(false),
                    disabled: connectorSelection !== undefined ? connectorSelection.length <= 0 : true
                },
                {
                    key: 'deleteConnector',
                    text: 'Delete',
                    iconProps: { iconName: 'Delete' },
                    onClick: () => window.confirm('This action will remove selected connector and all associated routing rules. Proceed with caution.') ? this._onDeleteConnector() : this._doNothing(),
                    disabled: selectedConnector ? selectedConnector.system : true
                }
            ]
        };

        var listData = [
            {
                listProps: {
                    checkButtonAriaLabel: 'Select item',
                    ariaLabelForSelectAllCheckbox: 'Select all items',
                    items: connectors,
                    columns: columns,
                    selectionMode: SelectionMode.single,
                    ariaLabelForGrid: 'Notification connector list, use arrows to navigate'
                },
                key: 'connectors'
            }
        ];

        var newConnectorDropdownOptions: IDropdownOption[] = [];
        for (const connectorDefinition of connectorDefinitions) {
            if (!connectorDefinition.unique ||
                (connectorDefinition.unique === true && connectors.find((c: IConnector) => c.type === connectorDefinition.id) === undefined))
                newConnectorDropdownOptions.push({
                    key: connectorDefinition.id,
                    text: connectorDefinition.name,
                    data: { icon: connectorDefinition.icon }
                });
        }
        
        return (
            <AdminLayout>
            <div className="container" style={{ maxWidth: '97%' }}>
                <div className="row">
                    <div className="col">
                        <M365Breadcrumb
                            items={[
                                { text: 'Home', key: 'home', href: '/admin' },
                                { text: 'Notifications', key: 'notifications', href: '/admin/notifications' },
                                { text: 'Connectors', key: 'connectors', isCurrentItem: true }
                            ]}
                            style={{ marginBottom: '16px' }}
                        />
                        <DetailPageHeader
                            title="Connectors"
                            description="Connect notification systems by defining and configuring the connections in the table below."
                        />
                    </div>
                </div>
                <div className="row">
                    <div className="col">
                        <CompositeList
                            detailsListProps={listData}
                            commandBarProps={commandBarProps}
                            defaultIsCompactMode={true}
                            onSelectionChanged={this._onSelectionChanged}
                            legacyInvokeAndSelect={true}
                            isEmpty={
                                listData === undefined ||
                                listData[0].listProps?.items === undefined ||
                                listData[0].listProps?.items?.length === 0 }
                            emptyStateProps={{
                                title: 'No connectors found',
                                body: 'Create at least one connector to enable notifications.',
                                imageProps: {
                                    src: '/images/empty.png',
                                    alt: 'No connectors found graphic'
                                },
                                illustrationSize: EmptyStateImageSize.Large,
                                actionBarProps: {
                                    primaryButtonProps: {
                                        text: 'New connector',
                                        onClick: () => alert('Create new rule event.')
                                    }
                                }
                            }}
                        />
                        <Dialog
                            title={selectedConnector ? selectedConnector.name : "New connector"}
                            isOpen={isEditConnectorPanelOpen}
                            onDismiss={this._onDismisEditConnectorPanel}
                            type={DialogType.largeHeader}
                            isBlocking={true}
                            minWidth={350}
                            // You MUST provide this prop! Otherwise screen readers will just say "button" with no label.
                            closeButtonAriaLabel="Close"
                        >
                            {editorPanelNewItemMode ? (<div>
                                <Dropdown
                                    title="Select connector type"
                                    options={newConnectorDropdownOptions}
                                    onRenderTitle={this._onRenderTitle}
                                    onRenderOption={this._onRenderOption}
                                    onChange={this._onChangeNewConnectorDropdown}
                                />
                            </div>) : (<div />)}
                            <TextField
                                label="Connector name"
                                onChange={this._onChangeNameTextField}
                                required={true}
                                defaultValue={newConnectorName ? newConnectorName : ''}
                            />
                            {(editorPanelNewItemMode && newItemConnectorType !== undefined) || (!editorPanelNewItemMode) ? (<div>
                                {selectedConnectorParameterDefinition.length > 0 ? (<div>
                                        {selectedConnectorParameterDefinition.map((d: IConnectorParameterDefinition) => (
                                            <TextField
                                                label={d.displayName}
                                                defaultValue={editorPanelNewItemMode ? '' : connectorEditorValues.find((c: IConnectorConfigurationValue) => c.name === d.name)?.value}
                                                onChange={this._onChangeTextField}
                                                accessKey={d.name}
                                                required={true}
                                            />
                                        ))
                                        }
                                        <br />
                                        <DialogFooter>
                                            <PrimaryButton onClick={this._onSaveConnectorConfig} text="Save" />
                                            <DefaultButton onClick={this._onDismisEditConnectorPanel} text="Cancel" />
                                        </DialogFooter>
                                    </div>
                                    ) : (
                                        <div>
                                            <Label>This connector doesn't contain any configuration options.</Label><br />
                                            <DialogFooter>
                                                <PrimaryButton onClick={this._onDismisEditConnectorPanel} text="Close" />
                                            </DialogFooter>
                                        </div>
                                    )
                                }</div>) : (
                                    <div>
                                        <DialogFooter>
                                            <PrimaryButton onClick={this._onDismisEditConnectorPanel} text="Close" />
                                        </DialogFooter>
                                    </div>
                                    )}
                        </Dialog>
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
        var connectors: IConnector[] = [];
        var connectorDefinitions: IConnectorDefinition[] = [];

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

                            this.setState({
                                connectors: connectors,
                                initialized: true
                            });
                        }).catch((err) => {
                            this.setState({
                                error: err.message
                            });
                        });
            })
            .then(() => {
                if (userHasRequiredRole)
                    fetch('/api/ConnectorDefinition?type=notificationManager', { headers: authResponse.idToken === "" ? {} : { 'Authorization': `Bearer ${authResponse.idToken}` } })
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
                                connectorDefinitions: connectorDefinitions,
                                initialized: true
                            });
                        }).catch((err) => {
                            this.setState({
                                error: err.message
                            });
                        });
            });
    }

    private _doNothing = (): void => {
        
    }
    private _onSelectionChanged = (selectionMap: ICompositeListSelectionMap, listKey: string) => {
        const activeSelection: ICompositeListSelectionMapItem | undefined = selectionMap.get(listKey);

        const newSelection: IObjectWithKey[] | undefined = activeSelection?.selection.getSelection();

        if (newSelection) {
            var selectedConnector: IConnector | undefined = newSelection === undefined ? undefined : newSelection[0] as IConnector;
            var parameterDefinition: IConnectorParameterDefinition[] = selectedConnector === undefined ? [] : selectedConnector.parameterDefinition.filter((a: IConnectorParameterDefinition) => a.parameterType === 'configuration');
            var selectedConfiguration: IConnectorConfigurationValue[] = selectedConnector === undefined ? [] : selectedConnector.configuration as IConnectorConfigurationValue[];
    
            this.setState({
                connectorSelection: newSelection,
                selectedConnector: selectedConnector,
                selectedConnectorParameterDefinition: parameterDefinition,
                selectedConnectorConfiguration: selectedConfiguration
            });
        } else {
            this.setState({
                connectorSelection: [],
                selectedConnector: undefined,
                selectedConnectorParameterDefinition: [],
                selectedConnectorConfiguration: []
            });
        }
    };

    _onOpenEditConnectorPanel(newItemMode: boolean) {
        var selectedConfiguration: IConnectorConfigurationValue[] = this.state.selectedConnector === undefined ? [] : this.state.selectedConnector.configuration as IConnectorConfigurationValue[];
        var editorConfiguration: IConnectorConfigurationValue[] = [];
        var connectorName = newItemMode ? '' : this.state.selectedConnector!.name;

        if (!newItemMode)
            for (const configEntry of selectedConfiguration) {
                editorConfiguration.push({
                    name: configEntry.name,
                    value: configEntry.value
                });
            }
        else {

        }

        this.setState({
            editorPanelNewItemMode: newItemMode,
            connectorEditorValues: editorConfiguration,
            newConnectorName: connectorName,
            isEditConnectorPanelOpen: true
        });
    }

    private _onDismisEditConnectorPanel = (): void => {
        this.setState({
            connectorEditorValues: [],
            isEditConnectorPanelOpen: false,
            editorPanelNewItemMode: false,
            newConnectorName: undefined,
            newItemConnectorType: undefined,
        });
    }

    private _onSaveConnectorConfig = (): void => {
        var newConfiguration: IConnectorConfigurationValue[] = [];

        for (const configEntry of this.state.connectorEditorValues) {
            newConfiguration.push({
                name: configEntry.name,
                value: configEntry.value
            });
        }

        var connectors = this.state.connectors;
        var connector = connectors.find((c: IConnector) => c.connectorId === this.state.selectedConnector?.connectorId);
        var params: RequestInit;

        if (this.state.editorPanelNewItemMode && this.state.newItemConnectorType !== undefined) {
            var guid: string = this._createGuid();
            var connectorDefinition: IConnectorDefinition | undefined = this.state.connectorDefinitions.find((cd: IConnectorDefinition) => cd.id === this.state.newItemConnectorType);

            connector = {
                key: guid,
                connectorId: guid,
                connectorType: 'notificationManager',
                configuration: newConfiguration,
                connectorTypeName: connectorDefinition ? connectorDefinition.name : '',
                icon: connectorDefinition ? connectorDefinition.icon : '',
                name: this.state.newConnectorName ? this.state.newConnectorName : 'Untitled',
                parameterDefinition: connectorDefinition ? connectorDefinition.parameters : [],
                system: connectorDefinition ? connectorDefinition.system : false,
                type: this.state.newItemConnectorType
            };
        }

        if (connector !== undefined) {
            connector.configuration = newConfiguration;
            connector.name = this.state.newConnectorName ? this.state.newConnectorName : connector.name;
            acquireAccessToken()
                .then((response) => {
                    params = {
                        headers: {
                            "Content-Type": "application/json charset=UTF-8",
                            "Authorization": "Bearer " + response.idToken
                        },
                        body: JSON.stringify(connector),
                        method: "POST"
                    };

                    fetch("/api/Connectors", params)
                        .then((response) => {
                            if (connectors.find((c: IConnector) => c.connectorId === connector!.connectorId) === undefined)
                                connectors.push(connector!);

                            connectors = _copyAndSort(connectors, 'name', false);

                            this.setState({
                                connectorEditorValues: [],
                                connectors: connectors,
                                selectedConnectorConfiguration: newConfiguration,
                                editorPanelNewItemMode: false,
                                newConnectorName: undefined,
                                newItemConnectorType: undefined,
                                isEditConnectorPanelOpen: false
                            });
                        })
                        .catch((err) => {
                            this.setState({
                                connectorEditorValues: [],
                                connectors: connectors,
                                selectedConnectorConfiguration: newConfiguration,
                                isEditConnectorPanelOpen: false,
                                editorPanelNewItemMode: false,
                                newConnectorName: undefined,
                                newItemConnectorType: undefined,
                                saveError: err.message
                            });
                        });
                })
                .catch((err) => {
                    this.setState({
                        connectorEditorValues: [],
                        connectors: connectors,
                        selectedConnectorConfiguration: newConfiguration,
                        editorPanelNewItemMode: false,
                        newConnectorName: undefined,
                        newItemConnectorType: undefined,
                        isEditConnectorPanelOpen: false,
                        saveError: "Authentication error. Details:<br/><br/>" + err.message
                    });
                });
        }
    }

    private _onDeleteConnector = (): void => {
        var connectors = this.state.connectors;
        var connector = connectors.find((c: IConnector) => c.connectorId === this.state.selectedConnector?.connectorId);
        var params: RequestInit;

        if (connector !== undefined) {
            acquireAccessToken()
                .then((response) => {
                    params = {
                        headers: {
                            "Content-Type": "application/json charset=UTF-8",
                            "Authorization": "Bearer " + response.idToken
                        },
                        body: JSON.stringify(connector),
                        method: "DELETE"
                    };

                    fetch("/api/Connectors", params)
                        .then((response) => {
                            if (connectors.find((c: IConnector) => c.connectorId === connector!.connectorId) !== undefined) 
                                connectors = connectors.filter((c: IConnector) => c.connectorId !== connector!.connectorId);

                            connectors = _copyAndSort(connectors, 'name', false);

                            this.setState({
                                connectorEditorValues: [],
                                connectors: connectors,
                                editorPanelNewItemMode: false,
                                newConnectorName: undefined,
                                newItemConnectorType: undefined,
                                isEditConnectorPanelOpen: false
                            });
                        })
                        .catch((err) => {
                            this.setState({
                                connectorEditorValues: [],
                                connectors: connectors,
                                isEditConnectorPanelOpen: false,
                                editorPanelNewItemMode: false,
                                newConnectorName: undefined,
                                newItemConnectorType: undefined,
                                saveError: err.message
                            });
                        });
                })
                .catch((err) => {
                    this.setState({
                        connectorEditorValues: [],
                        connectors: connectors,
                        editorPanelNewItemMode: false,
                        newConnectorName: undefined,
                        newItemConnectorType: undefined,
                        isEditConnectorPanelOpen: false,
                        saveError: "Authentication error. Details:<br/><br/>" + err.message
                    });
                });
        }
    }

    private _createGuid(): string {
        return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
            var r = Math.random() * 16 | 0, v = c === 'x' ? r : (r & 0x3 | 0x8);
            return v.toString(16);
        });
    }

    private _onChangeTextField = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
        var configEntry: IConnectorConfigurationValue | undefined = this.state.connectorEditorValues.find((c: IConnectorConfigurationValue) => c.name === ev.currentTarget.accessKey);
        if (configEntry !== undefined)
            configEntry.value = newValue ? newValue : "";
        else
            this.state.connectorEditorValues.push({
                name: ev.currentTarget.accessKey,
                value: newValue
            });
    };

    private _onChangeNameTextField = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
        this.setState({
            newConnectorName: newValue
        });
    };

    private _onRenderTitle = (options: ISelectableOption[] | undefined): JSX.Element => {
        const option = options  ? options[0] : undefined;

        return (
            <div>
                {option?.data && option?.data.icon && (
                    <img src={option.data.icon} aria-hidden="true" style={{ width: '32px', float: 'left' }} />
                )}
                <span>{option?.text}</span>
            </div>
        );
    };

    private _onRenderOption = (option: ISelectableOption | undefined): JSX.Element => {
        return (
            <div>
                {option?.data && option?.data.icon && (
                    <img src={option.data.icon} aria-hidden="true" style={{ width: '32px', marginTop: '-4px', float: 'left' }} />
                )}
                <span>{option?.text}</span>
            </div>
        );
    };

    private _onChangeNewConnectorDropdown = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number): void => {
        var parameterDefinition: IConnectorParameterDefinition[] = [];
        if (option !== undefined) {
            var selectedConnector = this.state.connectorDefinitions.find((cd: IConnectorDefinition) => cd.id === option.key);
            var parameterDefinition: IConnectorParameterDefinition[] = selectedConnector === undefined ? [] : selectedConnector.parameters.filter((a: IConnectorParameterDefinition) => a.parameterType === 'configuration');
            this.setState({
                connectorEditorValues: [],
                selectedConnectorParameterDefinition: parameterDefinition,
                newItemConnectorType: option.key ? option.key.toString() : ''
            });
        }
        else {
            this.setState({
                connectorEditorValues: [],
                selectedConnectorParameterDefinition: []
            });
        }
    };
}

function _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
    const key = columnKey as keyof T;
    return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
}