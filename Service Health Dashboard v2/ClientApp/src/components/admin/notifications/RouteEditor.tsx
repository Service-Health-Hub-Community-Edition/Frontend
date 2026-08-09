import * as React from 'react';
import {
    Dropdown, IDropdownOption, ISelectableOption,
    TextField, Checkbox, Text, ITextProps, Label,
    ICommandBarProps, IObjectWithKey, SelectionMode,
    DialogFooter, PrimaryButton, DefaultButton,
    Dialog, DialogType, Link, ILinkStyles, ILinkStyleProps,
    CommandBar, DetailsList, Selection
} from '@fluentui/react';
import { ICompositeListColumn } from '@m365-admin/composite-list';
import { EmptyStateImageSize, EmptyState } from '@m365-admin/empty-state';
import { AccessDenied } from "../../AccessDenied";
import { acquireAccessToken } from "../../../auth/AccessTokenHelper";
import { AuthenticationResult } from '@azure/msal-browser';
import { IConnector, IConnectorDefinition, IConnectorParameterDefinition, IConnectorConfigurationValue } from './Connectors';
import { IRoutingCondition, IRoutingElement, IComponent, IEntityProperty } from './Routing';
import { ConditionEditor, multiValueOperators, operatorList } from './ConditionEditor';

const pipeFabricStyles = (p: ILinkStyleProps): ILinkStyles => ({
    root: {
        textDecoration: 'none',
        color: p.theme.semanticColors.bodyText
    },
});

interface IRoutingState {
    route: IRoutingElement;
    component?: IComponent;
    connectors: IConnector[];
    connectorDefinitions: IConnectorDefinition[];
    connectorsDropdown: IDropdownOption[];
    iconDropdown: IDropdownOption[];
    languageOptions: IDropdownOption[];
    conditionSelection: IObjectWithKey[];
    selectedCondition: IRoutingCondition | undefined;
    initialized: boolean;
    accessGranted: boolean;
    error?: string;
    isEditConditionDialogOpen: boolean;
    selectedConnectorConfigurationDefinition: IConnectorParameterDefinition[];
}

var emptyRoute: IRoutingElement = {
    key: '',
    id: '',
    name: '',
    order: 10000,
    conditions: [],
    connector: undefined,
    connectorConfiguration: [],
    icon: '',
    stopProcessingOnMatch: false,
    hideWorkItemLink: false,
    language: undefined
}

export class RouteEditor extends React.Component<{
    route?: IRoutingElement,
    component?: IComponent,
    connectors: IConnector[],
    connectorDefinitions: IConnectorDefinition[],
    onSave?: (route: IRoutingElement) => void,
    onCancel?: () => void
}, IRoutingState> {  

    private selection = new Selection({
        onSelectionChanged: () => this._onConditionSelectionChanged()
    });

    constructor(props: {
        route?: IRoutingElement,
        component?: IComponent,
        connectors: IConnector[],
        connectorDefinitions: IConnectorDefinition[],
        onSave?: (route: IRoutingElement) => void,
        onCancel?: () => void
    }) {
        super(props);

        var route: IRoutingElement = emptyRoute;

        if (this.props.route)
            route = Object.assign({}, this.props.route);
        else {
            route = Object.assign({}, emptyRoute);
            route.conditions = [];
            route.connectorConfiguration = [];
        }

        this.state = {
            route: route,
            component: this.props.component,
            connectors: [],
            connectorDefinitions: [],
            connectorsDropdown: [],
            iconDropdown: [],
            languageOptions: [],
            conditionSelection: [],
            selectedCondition: undefined,
            initialized: false,
            accessGranted: false,
            error: undefined,
            isEditConditionDialogOpen: false,
            selectedConnectorConfigurationDefinition: []
        };

        this._onConditionSave = this._onConditionSave.bind(this);
        this._onConditionEditCancel = this._onConditionEditCancel.bind(this);
    }

    public render() {
        const {
            route, initialized, connectorsDropdown, iconDropdown, languageOptions, conditionSelection, isEditConditionDialogOpen,
            selectedConnectorConfigurationDefinition, component, selectedCondition
        } = this.state;

        if (!initialized)
            return (<div />);

        const columns: ICompositeListColumn[] = [
            {
                key: 'ruleProperty',
                name: 'Property',
                minWidth: 120,
                maxWidth: 120,
                fieldName: 'property',
                isResizable: false,
                isRowHeader: true,
                onRender: (item: IRoutingCondition) => {
                    return <Link onClick={(event) => {
                        event.preventDefault();
                        this._onOpenEditConditionDialog(false);
                    }} styles={pipeFabricStyles}>{
                            this.state.component?.entityProperties === undefined ? item.property : this.state.component.entityProperties.find((c: IEntityProperty) => c.entityProperty === item.property)?.displayName
                        }</Link>;
                        
                }
            },
            {
                key: 'operator',
                name: 'Operator',
                minWidth: 75,
                maxWidth: 75,
                isResizable: false,
                onRender: (item: IRoutingCondition) => {
                    return <div>
                        {operatorList.find((o) => o.key === item.operator) ? operatorList.find((o) => o.key === item.operator)!.text : item.operator}
                    </div>;
                }
            },
            {
                key: 'value',
                name: 'Value',
                minWidth: 120,
                maxWidth: 120,
                isResizable: false,
                onRender: (item: IRoutingCondition) => {
                    if (multiValueOperators.includes(item.operator))
                        return <div>{item.values ? item.values.join('; ') : ''}</div>
                    else
                        return <div>{item.value}</div>
                }
            },
            {
                key: 'logicOp',
                name: 'And/Or',
                minWidth: 60,
                maxWidth: 60,
                fieldName: 'logicOperator',
                isResizable: true
            }
        ];

        const commandBarProps: ICommandBarProps = {
            items: [
                {
                    key: 'newCondition',
                    text: 'New',
                    iconProps: { iconName: 'Add' },
                    onClick: () => this._onOpenEditConditionDialog(true),
                },
                {
                    key: 'editConnector',
                    text: 'Edit',
                    iconProps: { iconName: 'Edit' },
                    onClick: () => this._onOpenEditConditionDialog(false),
                    disabled: conditionSelection !== undefined ? conditionSelection.length <= 0 : true
                },
                {
                    key: 'deleteCondition',
                    text: 'Delete',
                    iconProps: { iconName: 'Delete' },
                    onClick: () => this._onDeleteCondition(),
                    disabled: conditionSelection !== undefined ? conditionSelection.length <= 0 : true
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

        var listData = [
            {
                listProps: {
                    checkButtonAriaLabel: 'Select item',
                    ariaLabelForSelectAllCheckbox: 'Select all items',
                    items: route.conditions,
                    columns: columns,
                    selectionMode: SelectionMode.single,
                    ariaLabelForGrid: 'Condition list, use arrows to navigate'
                },
                key: 'conditions'
            }

        ];

        return (
            <div className="container" style={{ maxWidth: '97%' }}>
                <div className="row">
                    <div className="col">
                        <TextField
                            label="Name"
                            onChange={this._onChangeTextField}
                            required={true}
                            defaultValue={route.name}
                            accessKey='sys_name'
                            underlined
                        />
                    </div>
                </div>
                <div className="row">
                    <div className="col">
                        <Dropdown
                            label="Icon"
                            options={iconDropdown}
                            defaultSelectedKey={route.icon}
                            onRenderTitle={this._onRenderTitle}
                            onRenderOption={this._onRenderOption}
                            onChange={this._onChangeIconDropdown}
                            accessKey='icon'
                        />
                    </div>
                    <div className="col">
                        <Dropdown
                            label="Language"
                            options={languageOptions}
                            defaultSelectedKey={route.language}
                            onChange={this._onChangeLanguageDropdown}
                            accessKey='language'
                        />
                    </div>
                </div>
                <div className="row" style={{ marginTop: '16px' }} >
                        <div className="col">
                            <CommandBar
                                items={commandBarProps.items}
                            />
                        <DetailsList
                            checkButtonAriaLabel='Select item'
                            ariaLabelForSelectAllCheckbox='Select all items'
                            items={route.conditions}
                            columns={columns}
                            selectionMode={SelectionMode.single}
                            selection={this.selection}
                            setKey='id'
                            ariaLabelForGrid='Condition list, use arrows to navigate'
                            compact={true}
                            />
                        </div>
                    </div>
                <div className="row" style={{ marginBottom: '16px' }}>
                        <div className="col">
                          {route.conditions === undefined ||
                            route.conditions.length === 0 ? (
                                <EmptyState
                                    title='No conditions found'
                                    body='This rule will match all items.'
                                    imageProps={{
                                        src: '/images/empty.png',
                                        alt: 'No conditions graphic'
                                    }}
                                    illustrationSize={EmptyStateImageSize.Small}

                                    actionBarProps={{
                                        primaryButtonProps: {
                                            text: 'New rule',
                                            onClick: () => this._onOpenEditConditionDialog(true)
                                        }
                                    }}
                                    styles={{ root: { justifyContent: '', position: 'relative' } }}
                                />) : ""}
                        </div>
                    </div>
                <div className="row" >
                    <div className="col">
                        <Label>Additional options</Label>
                        <Checkbox label="Stop processing on match" onChange={this._onStopProcessingChange} defaultChecked={route.stopProcessingOnMatch} />
                    </div>
                </div>
                <div className="row">
                    <div className="col">
                        <Checkbox label="Hide work item link" onChange={this._onHideTaskChange} defaultChecked={route.hideWorkItemLink} />
                    </div>
                </div>
                <div className="row">
                    <div className="col">
                        <Label>On match, send to</Label>
                    </div>
                </div>
                <div className="row">
                    <div className="col">
                        <Dropdown
                            label="Connector"
                            options={connectorsDropdown}
                            defaultSelectedKey={route.connector?.connectorId}
                            onRenderTitle={this._onRenderTitle}
                            onRenderOption={this._onRenderOption}
                            onChange={this._onChangeConnectorDropdown}
                            accessKey='connector'
                        />
                    </div>
                </div>
                <div className="row">
                    <div className="col">
                        <Label>Connector configuration</Label>
                    </div>
                </div>
                <div className="row">
                    {selectedConnectorConfigurationDefinition.map((d: IConnectorParameterDefinition) => (
                        <div className="col">
                            <TextField
                                label={d.displayName}
                                value={this._getConnectorParameterValue(d.name)}
                                onChange={this._onChangeTextField}
                                accessKey={d.name}
                                required={true}
                                underlined
                            />
                        </div>
                    ))
                    }
                </div>

                    <div className="row">
                        <div className="col">
                            <Dialog
                                title={"Condition"}
                                isOpen={isEditConditionDialogOpen}
                                onDismiss={this._onDismisEditConditionDialog}
                                type={DialogType.largeHeader}
                                isBlocking={true}
                                minWidth={350}
                                // You MUST provide this prop! Otherwise screen readers will just say "button" with no label.
                                closeButtonAriaLabel="Close"
                                styles={{
                                    main: {
                                        selectors: {
                                            ['@media (min-width: 480px)']: {
                                                width: 720,
                                                minWidth: 350,
                                                maxWidth: '720x'
                                            }
                                        }
                                    }
                                }}
                        >
                                <ConditionEditor condition={selectedCondition} component={component} onSave={this._onConditionSave} onCancel={this._onConditionEditCancel} />
                            </Dialog>
                        </div>
                    </div>

                <div className="row">
                    <div className="col">
                        <DialogFooter>
                            <PrimaryButton onClick={this._saveRoute} text="Save" />
                            <DefaultButton onClick={this._cancelEdit} text="Cancel" />
                        </DialogFooter>
                    </div>
                </div>
            </div>
        );
    }

    componentDidMount() {
        var connectorDropdownOptions: IDropdownOption[] = [];

        for (const connector of this.props.connectors) {
            connectorDropdownOptions.push({
                key: connector.connectorId,
                text: connector.name,
                data: {
                    icon: connector.icon
                }
            });
        }

        if (this.props.route !== undefined && this.props.route.conditions !== undefined) {
            var i: number = 0;

            for (var condition of this.props.route?.conditions) {
                condition.key = this._createGuid();
                if (condition.order === undefined) {
                    condition.order = i;
                }
                i++;
            }
        }

        var route: IRoutingElement = emptyRoute;

        if (this.props.route)
            route = Object.assign({}, this.props.route);
        else {
            route = Object.assign({}, emptyRoute);
            route.conditions = [];
            route.connectorConfiguration = [];
        }

        this.setState({
            route: route,
            connectors: this.props.connectors,
            connectorDefinitions: this.props.connectorDefinitions,
            connectorsDropdown: connectorDropdownOptions,
            initialized: true
        });

        if (this.props.route?.connector !== undefined)
            this._handleConnectorState(this.props.route.connector);

        this._getIcons();
        this._getLanguages();
    }

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

    private _getConnectorParameterValue = (name: string): string => {
        var value = this.state.route.connectorConfiguration.find((cfg: IConnectorConfigurationValue) => cfg.name === name)?.value;
        return value ? value : '';
    }

    private _onChangeConnectorDropdown = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number): void => {    
        if (option !== undefined) {
            this._handleConnectorChange(option.key.toString());
        } else {
            this.setState({
                selectedConnectorConfigurationDefinition: []
            });
        } 
    };

    private _onChangeLanguageDropdown = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number): void => {
        var route: IRoutingElement = this.state.route;
        route.language = option ? option.key.toString() : '';
        this.setState({
            route: route
        });
    };

    private _onChangeIconDropdown = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number): void => {
        var route: IRoutingElement = this.state.route;
        route.icon = option ? option.key.toString() : '';
        this.setState({
            route: route
        });
    };

    private _handleConnectorChange = (id: string): void => {
        var connector: IConnector | undefined = this.state.connectors.find((c) => c.connectorId === id);
        this._handleConnectorState(connector);
    }

    private _handleConnectorState = (connector?: IConnector): void => {
        var configurationDefinitionCollection: IConnectorParameterDefinition[] = [];

        if (connector?.parameterDefinition !== undefined)
            for (const configurationDefinition of connector?.parameterDefinition)
                if (configurationDefinition.parameterType === 'routing')
                    configurationDefinitionCollection.push(configurationDefinition);

        var route: IRoutingElement = this.state.route;
        route.connector = connector;

        this.setState({
            route: route,
            selectedConnectorConfigurationDefinition: configurationDefinitionCollection
        });
    }

    private _onChangeTextField = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
        var route: IRoutingElement = this.state.route;
        var value: string = newValue ? newValue : '';
        switch (ev.currentTarget.accessKey) {
            case 'sys_name':
                route.name = value;
                break;
            default:
                // custom connector config here
                var cfgValue: IConnectorConfigurationValue | undefined = route.connectorConfiguration.find((c: IConnectorConfigurationValue) => c.name === ev.currentTarget.accessKey);
                if (cfgValue === undefined)
                    route.connectorConfiguration.push({
                        name: ev.currentTarget.accessKey,
                        value: value
                    })
                else
                    cfgValue.value = value;

                break;
        }

        this.setState({
            route: route
        });
    };

    private _getIcons = (): void => {
        const requiredRoles: string[] = ['Admin'];
        var iconDropdownOptions: IDropdownOption[] = [];
        var userHasRequiredRole: boolean = false;
        var authResponse: AuthenticationResult;
        var key: number = 0;

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
                    fetch('/api/ImageStore?type=notificationIcon', { headers: authResponse.idToken === "" ? {} : { 'Authorization': `Bearer ${authResponse.idToken}` } })
                        .then((response) => {
                            return response.json();
                        })
                        .then((result) => {
                            for (const icon of result) {
                                iconDropdownOptions.push({
                                    key: 'imagestore://' + icon.name,
                                    text: icon.name,
                                    data: {
                                        icon: 'data:' + icon.format + ';base64,'+icon.content
                                    }
                                });
                                key++;
                            }

                            iconDropdownOptions = _copyAndSort(iconDropdownOptions, 'text', false);

                            this.setState({
                                iconDropdown: iconDropdownOptions
                            });
                        });
            }).catch((err) => {
                this.setState({
                    error: err.message
                });
            });
    }

    private _getLanguages = (): void => {
        var languageDropdownOptions: IDropdownOption[] = [];

        fetch('https://api.cognitive.microsofttranslator.com/languages?api-version=3.0')
            .then((response) => {
                return response.json();
            })
            .then((result) => {
                languageDropdownOptions.push({
                    key: '',
                    text: ''
                });

                for (const lang in result.translation) {
                    languageDropdownOptions.push({
                        key: lang,
                        text: result.translation[lang].name
                    });
                }

                languageDropdownOptions = languageDropdownOptions.sort((a, b) => 0 - (a.text > b.text ? -1 : 1));

                this.setState({
                    languageOptions: languageDropdownOptions
                });
            });
    }

    private _onStopProcessingChange = (ev?: React.FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean): void => {
        var route: IRoutingElement = this.state.route;
        route.stopProcessingOnMatch = checked ? checked : false;
        this.setState({
            route: route
        });
    }

    private _onHideTaskChange = (ev?: React.FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean): void => {
        var route: IRoutingElement = this.state.route;
        route.hideWorkItemLink = checked ? checked : false;
        this.setState({
            route: route
        });
    }

    private _createGuid(): string {
        return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
            var r = Math.random() * 16 | 0, v = c === 'x' ? r : (r & 0x3 | 0x8);
            return v.toString(16);
        });
    }

    private _onConditionSelectionChanged = () => {
        const items: IObjectWithKey[] = this.selection.getSelection();
        var selectedCondition: IRoutingCondition | undefined = items === undefined || items.length <= 0 ? undefined : items[0] as IRoutingCondition;

        if (selectedCondition) {
            this.setState({
                conditionSelection: items,
                selectedCondition: selectedCondition
            });
        } else {
            this.setState({
                conditionSelection: [],
                selectedCondition: undefined
            });
        }
    };

    _onOpenEditConditionDialog(newItemMode: boolean) {
        var selectedCondition: IRoutingCondition[] = this.state.conditionSelection === undefined ? [] : this.state.conditionSelection as IRoutingCondition[];
        var editorConfiguration: IRoutingCondition[] = [];

        if (newItemMode)
            this.setState({
                selectedCondition: undefined,
                conditionSelection: [],
                isEditConditionDialogOpen: true
            });
        else
            this.setState({
                isEditConditionDialogOpen: true
            });
    }

    private _onDismisEditConditionDialog = (): void => {
        this.setState({
            isEditConditionDialogOpen: false
        });
    }

    private _isMoveUpDownDisabled = (first: boolean): boolean => {
        if (this.state.route.conditions !== undefined && this.state.route.conditions.length > 0 &&
            this.state.conditionSelection !== undefined && this.state.conditionSelection.length > 0) {
            var orderedConditions = _copyAndSort(this.state.route.conditions, 'order', !first);
            var firstCondition = orderedConditions[0];
            var selection: IRoutingCondition = this.state.conditionSelection[0] as IRoutingCondition;
            return selection ? selection.order === firstCondition.order : false;
        } else {
            return true;
        }
    }

    private _move = (up: boolean): void => {
        if (this.state.conditionSelection !== undefined && this.state.conditionSelection.length > 0) {
            var selection: IRoutingCondition = this.state.conditionSelection[0] as IRoutingCondition;
            var selectedOrderValue: number = selection.order!;
            var currentCondition = this.state.route.conditions.find(element => element.order === selectedOrderValue);
            if (currentCondition !== undefined) {
                var index = this.state.route.conditions.indexOf(currentCondition!);
                var movingCondition: IRoutingCondition | undefined = undefined;
                if (up) {
                    if (index > 0)
                        movingCondition = this.state.route.conditions[index - 1];
                } else {
                    if (index < this.state.route.conditions.length - 1)
                        movingCondition = this.state.route.conditions[index + 1];
                }

                if (movingCondition !== undefined) {
                    currentCondition.order = movingCondition.order;
                    movingCondition.order = selectedOrderValue;
                    var route = this.state.route;

                    /* this._onSaveRoute(currentCondition);
                    this._onSaveRoute(movingCondition); */

                    route.conditions.sort((a: IRoutingCondition, b: IRoutingCondition) => a.order! > b.order! ? 1 : -1);

                    this.selection.setItems(route.conditions);

                    this.state.conditionSelection.map(cs => {
                        if (cs.key !== undefined && cs.key !== null)
                            this.selection.setKeySelected(cs.key as string, true, false)
                    });

                    this.setState({
                        route: route
                    });
                }
            }
        }
    }

    private _doNothing = (): void => {

    }

    private _onConditionSave(condition: IRoutingCondition): void {
        var route: IRoutingElement = this.state.route;
        var oldCondition: IRoutingCondition | undefined = route.conditions.find((c) => c.key === condition.key);
        if (oldCondition !== undefined) {
            oldCondition.logicOperator = condition.logicOperator;
            oldCondition.operator = condition.operator;
            oldCondition.order = condition.order;
            oldCondition.property = condition.property;
            oldCondition.value = condition.value;
            oldCondition.values = condition.values;
        } else {
            condition.key = this._createGuid();
            var nextOrderId: number = 0;
            var orderedConditions = _copyAndSort(this.state.route.conditions, 'order', true);
            if (orderedConditions !== undefined && orderedConditions.length > 0) {
                var firstCondition = orderedConditions[0];
                nextOrderId = firstCondition.order ? firstCondition.order + 1 : 1000;
            }
            condition.order = nextOrderId;
            route.conditions.push(condition);
        }

        this.setState({
            route: route,
            isEditConditionDialogOpen: false
        });
    }

    private _onConditionEditCancel(): void {
        this.setState({
            isEditConditionDialogOpen: false
        });
    }

    private _onDeleteCondition = (): void => {
        var route: IRoutingElement | undefined = this.state.route;

        if (route !== undefined &&
            this.state.selectedCondition !== undefined &&
            route.conditions.find((c: IRoutingCondition) => c.key === this.state.selectedCondition!.key) !== undefined) {
            route.conditions = route.conditions.filter((c: IRoutingCondition) => c.key !== this.state.selectedCondition!.key);

            this.setState({
                route: route,
                selectedCondition: undefined,
                conditionSelection: []
            });
        }
    }

    private _saveRoute = (): void => {
        this.props.onSave ? this.props.onSave(this.state.route) : this._doNothing();
    }

    private _cancelEdit = (): void => {
        this.props.onCancel ? this.props.onCancel() : this._doNothing();
    }
}

function _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
    const key = columnKey as keyof T;
    return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
}
