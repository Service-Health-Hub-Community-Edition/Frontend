import * as React from 'react';
import {
    SelectionMode, IObjectWithKey, ICommandBarProps,
    Panel, PanelType, Dialog, DialogType, TextField, Text, Link, ILinkStyleProps, ILinkStyles,
    PrimaryButton, DefaultButton, DialogFooter, Label, Dropdown, IDropdownOption, ISelectableOption,
    Icon, IBreadcrumbItem, Separator
} from '@fluentui/react';
import { DetailPageHeader } from '@m365-admin/detail-page';
import { IconAlert, IconAlertSize, IconAlertStatus } from '@m365-admin/icon-alert';
import { CompositeList, ICompositeListColumn, ICompositeListSelectionMap, ICompositeListSelectionMapItem } from '@m365-admin/composite-list';
import { EmptyStateImageSize } from '@m365-admin/empty-state';
import { M365Breadcrumb } from '@m365-admin/m365-breadcrumb';
import { IconSelector } from '../IconSelector';
import { AccessDenied } from "../../AccessDenied";
import { AdminLayout } from '../AdminLayout';
import { acquireAccessToken } from "../../../auth/AccessTokenHelper";
import { AuthenticationResult } from '@azure/msal-browser';

const buttonStyles = { root: { marginRight: 8 } };

interface IEditorValue {
    name: string;
    value: string;
}

export interface IAction {
    key: string;
    actionId: string;
    componentId: string;
    name: string;
    icon: string;
    typeId: string;
    type: any;
    configuration: any;
}

export interface IActionTypeParameter {
    internalName: string;
    displayName: string;
    type: string;
}

export interface IActionType {
    actionTypeId: string;
    name: string;
    icon: string;
    parameters: IActionTypeParameter[];
}

interface IActionsState {
    entityId: string | null;
    entityName: string;
    actions: IAction[];
    actionSelection: IObjectWithKey[];
    actionEditorValues: any;
    selectedAction: IAction | undefined;
    editorActionType: IActionType | undefined;
    selectedActionConfiguration: any;
    actionTypeDefinitions: IActionType[];
    isEditActionPanelOpen: boolean;
    editorPanelNewItemMode: boolean;
    newItemActionType?: string;
    newActionName?: string;
    newIcon?: string;
    selectedId?: string;
    initialized: boolean;
    accessGranted: boolean;
    error?: string;
    saveError?: string;
}

export class Actions extends React.Component<{}, IActionsState> {

    constructor(props: {}) {
        super(props);

        const queryParams = new URLSearchParams(window.location.search);
        const entityId = queryParams.get('entityId');

        this.state = {
            entityId: entityId,
            entityName: "",
            actions: [],
            actionSelection: [],
            actionEditorValues: undefined,
            selectedAction: undefined,
            selectedActionConfiguration: [],
            editorActionType: undefined,
            actionTypeDefinitions: [],
            isEditActionPanelOpen: false,
            editorPanelNewItemMode: false,
            newItemActionType: undefined,
            newActionName: undefined,
            newIcon: '',
            selectedId: undefined,
            initialized: false,
            accessGranted: false,
            error: undefined,
            saveError: undefined
        };
    }

    public render() {
        const {
            entityId, entityName, actions, initialized, actionSelection, selectedAction,
            actionEditorValues, isEditActionPanelOpen, actionTypeDefinitions,
            editorPanelNewItemMode, newActionName, newItemActionType, editorActionType,
            newIcon
        } = this.state;

        if (!initialized)
            return (<div />);

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
                onRender: (item: IAction) => {
                    return <div>
                        <Icon iconName={item.icon} />
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
                onRender: (item: IAction) => {
                    return <Link onClick={(event) => {
                        event.preventDefault();
                        this._onOpenEditActionPanel(false);
                    }} styles={pipeFabricStyles}>{item.name}</Link >;
                }
            },
            {
                key: 'type',
                name: 'Type',
                minWidth: 280,
                maxWidth: 280,
                fieldName: 'type',
                isResizable: false,
                isRowHeader: true,
                onRender: (item: IAction) => {
                    return <div>
                        {item.type ? item.type['name'] : "Unknown"}
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
                    onClick: () => this._onOpenEditActionPanel(true),
                },
                {
                    key: 'editConnector',
                    text: 'Edit',
                    iconProps: { iconName: 'Edit' },
                    onClick: () => this._onOpenEditActionPanel(false),
                    disabled: actionSelection !== undefined ? actionSelection.length <= 0 : true
                },
                {
                    key: 'deleteConnector',
                    text: 'Delete',
                    iconProps: { iconName: 'Delete' },
                    onClick: () => window.confirm('This action will remove selected action. Proceed with caution.') ? this._onDeleteAction() : this._doNothing(),
                    disabled: selectedAction === undefined || selectedAction === null
                }
            ]
        };

        var listData = [
            {
                listProps: {
                    checkButtonAriaLabel: 'Select item',
                    ariaLabelForSelectAllCheckbox: 'Select all items',
                    items: actions,
                    columns: columns,
                    selectionMode: SelectionMode.single,
                    ariaLabelForGrid: 'Custom action list, use arrows to navigate'
                },
                key: 'actions'
            }
        ];

        var newActionDropdownOptions: IDropdownOption[] = [];
        for (const actionTypeDefinition of actionTypeDefinitions) {
            newActionDropdownOptions.push({
                key: actionTypeDefinition.actionTypeId,
                text: actionTypeDefinition.name,
                data: { icon: actionTypeDefinition.icon }
            });
        }

        const breadcrumbItems: IBreadcrumbItem[] = entityId !== null && entityName !== "" ?
            [
                { text: 'Home', key: 'home', href: '/admin' },
                { text: 'Integration', key: 'integration', href: '/admin/customActions/components' },
                { text: entityName, key: entityId! },
                { text: 'Actions', key: 'actions', isCurrentItem: true }
            ] :
            [
                { text: 'Home', key: 'home', href: '/admin' },
                { text: 'Integration', key: 'integration', href: '/admin/customActions/components' },
                { text: 'Actions', key: 'actions', isCurrentItem: true }
            ];


        return (
            <AdminLayout>
                <div className="container" style={{ maxWidth: '97%' }}>
                    <div className="row">
                        <div className="col">
                            <M365Breadcrumb
                                items={breadcrumbItems}
                                style={{ marginBottom: '16px' }}
                            />
                            <DetailPageHeader
                                title="Actions"
                                description="Customize integration by defining and configuring the actions in the table below."
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
                                    listData[0].listProps?.items?.length === 0}
                                emptyStateProps={{
                                    title: 'No actions found',
                                    body: 'Create at least one action to enable integration with 3rd party systems.',
                                    imageProps: {
                                        src: '/images/empty.png',
                                        alt: 'No actions found graphic'
                                    },
                                    illustrationSize: EmptyStateImageSize.Large,
                                    actionBarProps: {
                                        primaryButtonProps: {
                                            text: 'New action',
                                            onClick: () => alert('Create new rule event.')
                                        }
                                    }
                                }}
                            />
                            <Panel
                                headerText={selectedAction ? selectedAction.name : 'New action'}
                                isOpen={isEditActionPanelOpen}
                                onDismiss={this._onDismisEditActionPanel}
                                type={PanelType.medium}
                                isBlocking={true}
                                // You MUST provide this prop! Otherwise screen readers will just say "button" with no label.
                                closeButtonAriaLabel="Close"
                                isFooterAtBottom={true}
                                allowTouchBodyScroll={true}
                                onRenderFooterContent={() => (
                                    <div>
                                        <PrimaryButton onClick={this._onSaveAction} styles={buttonStyles}>Save</PrimaryButton> 
                                        <DefaultButton onClick={this._onDismisEditActionPanel} styles={buttonStyles}>Cancel</DefaultButton>
                                    </div>
                                    )}
                            >
                                <div className="container" style={{ maxWidth: '97%' }}>
                                    <div className="row">
                                        <div className="col">
                                            {editorPanelNewItemMode ? (<div>
                                                <Dropdown
                                                    label="Action"
                                                    required={true}
                                                    title="Select action type"
                                                    options={newActionDropdownOptions}
                                                    onRenderTitle={this._onRenderTitle}
                                                    onRenderOption={this._onRenderOption}
                                                    onChange={this._onChangeNewActionTypeDropdown}
                                                />
                                            </div>) : (<div />)}
                                        </div>
                                    </div>

                                    <div className="row">
                                        <div className="col">
                                            <TextField
                                                label="Action name"
                                                onChange={this._onChangeNameTextField}
                                                required={true}
                                                defaultValue={newActionName ? newActionName : ''}
                                            />
                                        </div>
                                        <div className="col">
                                            <IconSelector selectedIcon={newIcon} onChange={this._onChangeIcon.bind(this)} />
                                        </div>
                                    </div>

                                    <div className="row">
                                        <div className="col">
                                            {editorActionType ? (
                                                <div style={{ marginTop: '16px', marginBottom: '8px' }} >
                                                    <Separator />
                                                    <Text variant={'mediumPlus'}>{editorActionType.name} action configuration</Text>
                                                    </div>) : (<div />)}
                                        </div>
                                    </div>

                                    <div className="row">
                                        <div className="col">
                                            {editorActionType?.parameters?.map((d: IActionTypeParameter) => (
                                                <TextField
                                                    label={d.displayName}
                                                    defaultValue={editorPanelNewItemMode || actionEditorValues === undefined || actionEditorValues === null ? '' : actionEditorValues.find((c: IEditorValue) => c.name === d.internalName)?.value}
                                                    onChange={this._onChangeTextField}
                                                    accessKey={d.internalName}
                                                    required={true}
                                                />
                                            ))
                                            }
                                        </div>
                                    </div>
                                </div>      
                            </Panel>
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
        var actions: IAction[] = [];
        var actionTypes: IActionType[] = [];

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
                    fetch('/api/Components?id=' + this.state.entityId, { headers: authResponse.idToken === "" ? {} : { 'Authorization': `Bearer ${authResponse.idToken}` } })
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
                            if (result.length >= 1)
                                this.setState({
                                    entityName: result[0].name
                                });
                            else
                                this.setState({
                                    entityName: ""
                                });
                        }).catch((err) => {
                            this.setState({
                                error: err.message
                            });
                        });
            })
            .then(() => {
                if (userHasRequiredRole)
                    fetch('/api/customAction/actionType', { headers: authResponse.idToken === "" ? {} : { 'Authorization': `Bearer ${authResponse.idToken}` } })
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
                            for (const actionType of result) {
                                actionTypes.push({
                                    actionTypeId: actionType.id,
                                    icon: actionType.icon,
                                    name: actionType.name,
                                    parameters: actionType.parameters
                                });
                            }

                            this.setState({
                                actionTypeDefinitions: actionTypes,
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
                    fetch('/api/customAction?componentId=' + this.state.entityId, { headers: authResponse.idToken === "" ? {} : { 'Authorization': `Bearer ${authResponse.idToken}` } })
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
                            for (const action of result) {
                                actions.push({
                                    key: action.id,
                                    actionId: action.id,
                                    componentId: this.state.entityId ? this.state.entityId : '',
                                    name: action.name,
                                    typeId: action.type?.id,
                                    type: action.type,
                                    icon: action.icon,
                                    configuration: action.configuration
                                });
                            }

                            actions = _copyAndSort(actions, 'name', false);

                            this.setState({
                                actions: actions,
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
            var selectedAction: IAction | undefined = newSelection === undefined ? undefined : newSelection[0] as IAction;
            var selectedActionConfiguration: any = selectedAction === undefined ? {} : selectedAction.configuration;

            this.setState({
                actionSelection: newSelection,
                selectedAction: selectedAction,
                selectedActionConfiguration: selectedActionConfiguration
            });
        } else {
            this.setState({
                actionSelection: [],
                selectedAction: undefined,
                selectedActionConfiguration: {}
            });
        }
    };

    _onOpenEditActionPanel(newItemMode: boolean) {
        var selectedConfiguration: any = this.state.selectedAction === undefined ? undefined : this.state.selectedAction.configuration;
        var editorConfiguration: IEditorValue[] = [];
        var actionName = newItemMode ? '' : this.state.selectedAction!.name;
        var actionIcon = newItemMode ? '' : this.state.selectedAction!.icon;
        var editorActionType = newItemMode ? undefined : this.state.selectedAction?.type;
        if (!newItemMode)
            for (const configKeyName of Object.keys(selectedConfiguration)) {
                editorConfiguration.push({
                    name: configKeyName,
                    value: selectedConfiguration[configKeyName]
                });
            }
        else {

        }

        this.setState({
            editorPanelNewItemMode: newItemMode,
            actionEditorValues: editorConfiguration,
            editorActionType: editorActionType,
            newActionName: actionName,
            newIcon: actionIcon,
            isEditActionPanelOpen: true
        });

    }

    private _onDismisEditActionPanel = (): void => {
        this.setState({
            actionEditorValues: [],
            isEditActionPanelOpen: false,
            editorPanelNewItemMode: false,
            editorActionType: undefined,
            newActionName: undefined,
            newItemActionType: undefined,
        });
    }

    private _onSaveAction = (): void => {
        var newConfiguration: any = { };

        for (const configEntry of this.state.actionEditorValues) {
            newConfiguration[configEntry.name] = configEntry.value;
        }

        var actions = this.state.actions;
        var action = actions.find((a: IAction) => a.actionId === this.state.selectedAction?.actionId);
        var params: RequestInit;

        var actionBodyObj: any = undefined;

        if (this.state.editorPanelNewItemMode && this.state.newItemActionType !== undefined) {
            var guid: string = this._createGuid();
            var actionType: IActionType | undefined = this.state.actionTypeDefinitions.find((at: IActionType) => at.actionTypeId === this.state.newItemActionType);

            actionBodyObj = {
                id: guid,
                component: this.state.entityId,
                name: this.state.newActionName,
                icon: this.state.newIcon,
                type: this.state.newItemActionType,
                configuration: newConfiguration
            };
        } else {
            actionBodyObj = {
                id: action?.actionId,
                component: this.state.entityId,
                name: this.state.newActionName ? this.state.newActionName : action?.name,
                icon: this.state.newIcon ? this.state.newIcon : action?.icon,
                type: this.state.newItemActionType ? this.state.newItemActionType : action?.typeId,
                configuration: newConfiguration
            };
        }

        if (actionBodyObj !== undefined) {
            acquireAccessToken()
                .then((response) => {
                    params = {
                        headers: {
                            "Content-Type": "application/json charset=UTF-8",
                            "Authorization": "Bearer " + response.idToken
                        },
                        body: JSON.stringify(actionBodyObj),
                        method: "POST"
                    };

                    fetch("/api/customAction", params)
                        .then((response) => {
                            var updatedAction: IAction | undefined = actions.find((a: IAction) => a.actionId === actionBodyObj.id);
                            if (updatedAction === undefined) {
                                const newAction: IAction = {
                                    actionId: actionBodyObj.id,
                                    componentId: this.state.entityId ? this.state.entityId : '',
                                    icon: actionBodyObj.icon,
                                    key: actionBodyObj.id,
                                    name: actionBodyObj.name,
                                    typeId: actionBodyObj.type,
                                    type: this.state.actionTypeDefinitions.find((at: IActionType) => at.actionTypeId === actionBodyObj.type),
                                    configuration: newConfiguration
                                }

                                actions.push(newAction);
                            } else {
                                updatedAction.actionId = actionBodyObj.id;
                                updatedAction.componentId = this.state.entityId ? this.state.entityId : '';
                                updatedAction.icon = actionBodyObj.icon;
                                updatedAction.key = actionBodyObj.id;
                                updatedAction.name = actionBodyObj.name;
                                updatedAction.typeId = actionBodyObj.type;
                                updatedAction.type = this.state.actionTypeDefinitions.find((at: IActionType) => at.actionTypeId === actionBodyObj.type);
                                updatedAction.configuration = newConfiguration;
                            }

                            actions = _copyAndSort(actions, 'name', false);

                            this.setState({
                                actionEditorValues: [],
                                actions: actions,
                                selectedActionConfiguration: newConfiguration,
                                editorPanelNewItemMode: false,
                                newActionName: undefined,
                                newIcon: undefined,
                                newItemActionType: undefined,
                                isEditActionPanelOpen: false
                            });
                        })
                        .catch((err) => {
                            this.setState({
                                actionEditorValues: [],
                                actions: actions,
                                selectedActionConfiguration: newConfiguration,
                                editorPanelNewItemMode: false,
                                newActionName: undefined,
                                newIcon: undefined,
                                newItemActionType: undefined,
                                isEditActionPanelOpen: false,
                                saveError: err.message
                            });
                        });
                })
                .catch((err) => {
                    this.setState({
                        actionEditorValues: [],
                        actions: actions,
                        selectedActionConfiguration: newConfiguration,
                        editorPanelNewItemMode: false,
                        newActionName: undefined,
                        newIcon: undefined,
                        newItemActionType: undefined,
                        isEditActionPanelOpen: false,
                        saveError: "Authentication error. Details:<br/><br/>" + err.message
                    });
                });
        }
    }

    private _onDeleteAction = (): void => {
        var actions = this.state.actions;
        var action = actions.find((a: IAction) => a.actionId === this.state.selectedAction?.actionId);
        var params: RequestInit;

        if (action !== undefined) {
            acquireAccessToken()
                .then((response) => {

                    const actionBodyObj = {
                        id: action?.actionId,
                        component: action?.componentId,
                        name: action?.name,
                        icon: action?.icon,
                        type: action?.typeId,
                        configuration: action?.configuration
                    };

                    params = {
                        headers: {
                            "Content-Type": "application/json charset=UTF-8",
                            "Authorization": "Bearer " + response.idToken
                        },
                        body: JSON.stringify(actionBodyObj),
                        method: "DELETE"
                    };

                    fetch("/api/customAction", params)
                        .then((response) => {
                            if (actions.find((a: IAction) => a.actionId === action!.actionId) !== undefined)
                                actions = actions.filter((a: IAction) => a.actionId !== action!.actionId);

                            actions = _copyAndSort(actions, 'name', false);

                            this.setState({
                                actionEditorValues: [],
                                actions: actions,
                                editorPanelNewItemMode: false,
                                newActionName: undefined,
                                newIcon: '',
                                newItemActionType: undefined,
                                isEditActionPanelOpen: false
                            });
                        })
                        .catch((err) => {
                            this.setState({
                                actionEditorValues: [],
                                actions: actions,
                                editorPanelNewItemMode: false,
                                newActionName: undefined,
                                newIcon: '',
                                newItemActionType: undefined,
                                isEditActionPanelOpen: false,
                                saveError: err.message
                            });
                        });
                })
                .catch((err) => {
                    this.setState({
                        actionEditorValues: [],
                        actions: actions,
                        editorPanelNewItemMode: false,
                        newActionName: undefined,
                        newIcon: '',
                        newItemActionType: undefined,
                        isEditActionPanelOpen: false,
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

    _onChangeIcon = (icon: string): void => {
        this.setState({
            newIcon: icon
        });
    }

    private _onChangeTextField = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
        var configEntry: IEditorValue | undefined = this.state.actionEditorValues.find((c: IEditorValue) => c.name === ev.currentTarget.accessKey);
        if (configEntry !== undefined)
            configEntry.value = newValue ? newValue : "";
        else
            this.state.actionEditorValues.push({
                name: ev.currentTarget.accessKey,
                value: newValue
            });
    };

    private _onChangeNameTextField = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
        this.setState({
            newActionName: newValue
        });
    };

    private _onRenderTitle = (options: ISelectableOption[] | undefined): JSX.Element => {
        const option = options ? options[0] : undefined;

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

    private _onChangeNewActionTypeDropdown = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number): void => {
        if (option !== undefined) {
            var selectedActionType = this.state.actionTypeDefinitions.find((at: IActionType) => at.actionTypeId === option.key);
            this.setState({
                actionEditorValues: [],
                editorActionType: selectedActionType,
                newItemActionType: option.key ? option.key.toString() : ''
            });
        }
        else {
            this.setState({
                actionEditorValues: [],
                editorActionType: selectedActionType
            });
        }
    };
}

function _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
    const key = columnKey as keyof T;
    return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
}