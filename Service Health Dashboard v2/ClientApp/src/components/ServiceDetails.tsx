import * as React from 'react';
import { FontIcon, FontSizes, IIconProps, ILinkStyleProps, ILinkStyles, Link, Text, TextField, SearchBox } from '@fluentui/react';
import { Fabric } from '@fluentui/react';
import {
    DetailsList,
    DetailsListLayoutMode,
    Selection,
    SelectionMode,
    IColumn,
} from '@fluentui/react';
import { Panel, PanelType } from '@fluentui/react';
import { mergeStyleSets } from '@fluentui/react';
import { mergeStyles } from '@fluentui/react';
import { Dropdown, DropdownMenuItemType, IDropdownOption, IDropdownStyles } from '@fluentui/react';
import { DefaultButton } from '@fluentui/react';
import { CommandBar, ICommandBarItemProps } from '@fluentui/react';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { IBreadcrumbItem } from '@fluentui/react';
import { M365Breadcrumb } from '@m365-admin/m365-breadcrumb';
import { DetailPageHeader } from '@m365-admin/detail-page';
import { IncidentDetails } from './IncidentDetails';
import { IncidentCalendar } from './IncidentCalendar';
import { MessageBar, MessageBarType } from '@fluentui/react';
import { acquireAccessToken, accessControl } from "../auth/AccessTokenHelper";
import { AuthenticationResult } from '@azure/msal-common';
import { IServiceStatistics } from './IServiceStatistics';
import { checkInTeams } from './auth/detectTeams';
import { AccessDenied } from "./AccessDenied";

const dropdownStyles: Partial<IDropdownStyles> = { dropdown: { width: 300 } };

const classNames = mergeStyleSets({
    fileIconHeaderIcon: {
        padding: 0,
        fontSize: '16px',
    },
    fileIconCell: {
        textAlign: 'center',
        selectors: {
            '&:before': {
                content: '.',
                display: 'inline-block',
                verticalAlign: 'middle',
                height: '100%',
                width: '0px',
                visibility: 'hidden',
            },
        },
    },
    fileIconImg: {
        verticalAlign: 'middle',
        maxHeight: '16px',
        maxWidth: '16px',
    },
    controlWrapper: {
        display: 'flex',
        flexWrap: 'wrap',
    },
    exampleToggle: {
        display: 'inline-block',
        marginBottom: '10px',
        marginRight: '30px',
    },
    selectionDetails: {
        marginBottom: '20px',
    }
});
const controlStyles = {
    root: {
        margin: '0 30px 20px 0',
        maxWidth: '300px',
    },
};

const iconClass = mergeStyles({
});

const iconClasses = mergeStyleSets({
    ok: [{ color: 'green' }, iconClass],
    warning: [{ color: 'orange' }, iconClass],
    error: [{ color: 'red' }, iconClass],
});

export interface IServiceListState {
    columns: IColumn[];
    items: IServiceStatus[];
    selectionDetails: string;
    isModalSelection: boolean;
    isCompactMode: boolean;
    announcedMessage?: string;
    impactedServicesOnly?: boolean;
    textFilter?: string;
    isDataLoaded: boolean;
    isPanelOpen: boolean;
    selectedService: string;
    selectedFeatureListState: IFeatureStatus[];
    selectedServiceIncidents: IDropdownOption[];
    selectedIncidentItem?: IDropdownOption;
    selectedIncidentItemText?: string;
    isCalenderPanelOpen: boolean;
    inTeams: boolean;
    error?: string;
    accessGranted?: boolean;
}

export interface IFeatureStatus {
    feature: string;
    status: string;
}

export interface IServiceStatus {
    key: string;
    name: string;
    value: string;
    service: string;
    serviceStatus: string;
    issues: any;
    featureStatus: IFeatureStatus[];
}

enum ServiceHealthState {
    Ok = 0,
    Warning = 1,
    Error = 2
}

const _featureListcolumns = [
    {
        key: 'flcServiceStateIcon',
        name: 'ServiceState',
        className: classNames.fileIconCell,
        iconClassName: classNames.fileIconHeaderIcon,
        isIconOnly: true,
        fieldName: 'name',
        minWidth: 16,
        maxWidth: 16,
        onRender: (item: IFeatureStatus) => {
            var serviceHealthState = _getServiceHealthState(item.status);
            switch (serviceHealthState) {
                case ServiceHealthState.Error:
                    return <FontIcon iconName="StatusErrorFull" className={iconClasses.error} />
                case ServiceHealthState.Warning:
                    return <FontIcon iconName="AlertSolid" className={iconClasses.warning} />
                default:
                    return <FontIcon iconName="CompletedSolid" className={iconClasses.ok} />
            }
        },
    },
    {
        key: 'flcFeature',
        name: 'Feature',
        fieldName: 'feature',
        minWidth: 200,
        maxWidth: 250,
        isResizable: false
    },
    {
        key: 'flcState',
        name: 'State',
        fieldName: 'status',
        minWidth: 50,
        maxWidth: 150,
        isResizable: false
    },
];

const searchIcon: IIconProps = { iconName: 'Search' };

export class ServiceDetailsList extends React.Component<{}, IServiceListState> {
    private _selection: Selection;
    private _allItems: IServiceStatus[];
    private _breadcrumbItems: IBreadcrumbItem[] = [];  

    constructor(props: {}) {
        super(props);

        const columns: IColumn[] = [
            {
                key: 'column0',
                name: 'ServiceState',
                className: classNames.fileIconCell,
                iconClassName: classNames.fileIconHeaderIcon,
                isIconOnly: true,
                fieldName: 'name',
                minWidth: 16,
                maxWidth: 16,
                onRender: (item: IServiceStatus) => {
                    var serviceHealthState = _getServiceHealthState(item.serviceStatus);
                    switch (serviceHealthState) {
                        case ServiceHealthState.Error:
                            return <FontIcon iconName="StatusErrorFull" className={iconClasses.error} />
                        case ServiceHealthState.Warning:
                            return <FontIcon iconName="AlertSolid" className={iconClasses.warning} />
                        default:
                            if (item.issues !== undefined && item.issues.length > 0)
                                return <FontIcon iconName="AlertSolid" className={iconClasses.warning} />
                            else
                                return <FontIcon iconName="CompletedSolid" className={iconClasses.ok} />
                    }
                },
            },
            {
                key: 'column1',
                name: 'Service',
                fieldName: 'service',
                minWidth: 260,
                maxWidth: 300,
                isRowHeader: true,
                isResizable: true,
                isSorted: true,
                isSortedDescending: false,
                sortAscendingAriaLabel: 'Sorted A to Z',
                sortDescendingAriaLabel: 'Sorted Z to A',
                onColumnClick: this._onColumnClick,
                onRender: (item: IServiceStatus) => {
                    return <Link onClick={(event) => {
                        event.preventDefault();
                        this._onOpenPanel(item.service);
                    }} styles={pipeFabricStyles} > {item.service}</Link >;
                },
                data: 'string',
                isPadded: true,
            },
            {
                key: 'column2',
                name: 'Status',
                fieldName: 'statusDisplayName',
                minWidth: 180,
                maxWidth: 200,
                isResizable: true,
                isCollapsible: false,
                data: 'string',
                onColumnClick: this._onColumnClick,
                onRender: (item: IServiceStatus) => {
                    return <span>{item.serviceStatus}</span>;
                },
                isPadded: true,
            },
            {
                key: 'column3',
                name: 'Incidents',
                fieldName: 'incidents',
                minWidth: 60,
                maxWidth: 80,
                isResizable: true,
                isCollapsible: true,
                data: 'number',
                onColumnClick: this._onColumnClick,
                onRender: (item: IServiceStatus) => {
                    if (item.issues !== undefined && item.issues.length > 0) {
                        return <span>{this._getIncidentsGraphicalRepresentation(item.issues.length)}</span>;
                    }
                },
                isPadded: true,
            }
        ];

        this._allItems = [];

        this._selection = new Selection({
            onSelectionChanged: () => {
                this.setState({
                    selectionDetails: this._getSelectionDetails(),
                });
            },
        });

        this.state = {
            items: this._allItems,
            columns: columns,
            selectionDetails: this._getSelectionDetails(),
            isModalSelection: false,
            isCompactMode: false,
            announcedMessage: undefined,
            impactedServicesOnly: false,
            textFilter: "",
            isDataLoaded: false,
            isPanelOpen: false,
            selectedService: "",
            selectedFeatureListState: [],
            selectedServiceIncidents: [],
            selectedIncidentItem: undefined,
            selectedIncidentItemText: "",
            isCalenderPanelOpen: false,
            inTeams: checkInTeams(),
            error: undefined,
            accessGranted: undefined
        };


        this._breadcrumbItems = this.state.inTeams ? [
            { text: 'Status', key: 'status', isCurrentItem: false },
            { text: 'Service status', key: 'svcStatus', isCurrentItem: true }
        ] : 
            [
                { text: 'Home', key: 'home', isCurrentItem: false, href: '/' },
                { text: 'Status', key: 'status', isCurrentItem: false },
                { text: 'Service status', key: 'svcStatus', isCurrentItem: true }
            ];

        const pipeFabricStyles = (p: ILinkStyleProps): ILinkStyles => ({
            root: {
                textDecoration: 'none',
                color: p.theme.semanticColors.bodyText,
                fontWeight: '600',
                fontSize: p.theme.fonts.medium.fontSize,
            },
        });
    }

    public render() {
        const {
            columns, impactedServicesOnly, items,
            isDataLoaded, isPanelOpen, selectedService,
            selectedFeatureListState, selectedServiceIncidents, selectedIncidentItem,
            isCalenderPanelOpen, accessGranted
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

        const _commandBarItems: ICommandBarItemProps[] = [
            {
                key: 'openCalendar',
                text: 'Incident calendar',
                cacheKey: 'cacheIncidentCalendar', // changing this key will invalidate this item's cache
                iconProps: { iconName: 'Calendar' },
                onClick: () => this._onOpenCalenderPanel()
            }
        ];

        const _commandBarFarItems: ICommandBarItemProps[] = [
            {
                key: 'itemsCount',
                onRender: () => {
                    return <div style={{ paddingTop: 12, paddingRight:8 }}><Text>{items.length} item{items.length !== 1 ? "s" : ""}</Text></div>;
                },
            },
            {
                key: 'searchBox',
                onRender: () => {
                    return <div style={{ paddingLeft: 8, paddingRight: 8 }}><SearchBox placeholder="Search" iconProps={searchIcon} underlined={true} onChange={this._onChangeSearchText} /></div>;
                },
            },
            {
                key: 'servicesFilter',
                text: impactedServicesOnly ? 'Impacted services' : 'All services',
                iconProps: { iconName: 'List' },
                subMenuProps: {
                    items: [
                        {
                            key: 'allServices',
                            text: 'All services',
                            iconProps: { iconName: 'AllApps' },
                            checked: !impactedServicesOnly,
                            onClick: () => this._onImpactedServicesFilterChange(false)
                        },
                        {
                            key: 'impactedServices',
                            text: 'Impacted services',
                            iconProps: { iconName: 'AlertSolid', className: iconClasses.warning },
                            checked: impactedServicesOnly,
                            onClick: () => this._onImpactedServicesFilterChange(true)
                        }
                    ]
                },
            }
        ];

        return (
            <Fabric>
                <div className="container">
                    <div className="row">
                        <div className="col">
                            <M365Breadcrumb
                                items={this._breadcrumbItems}
                                maxDisplayedItems={10}
                                ariaLabel="Breadcrumb with items rendered as buttons"
                                overflowAriaLabel="More links"
                            />
                            <DetailPageHeader
                                title="Service status"
                                description="View the health status of all services that are available with your current subscriptions. View info about the history of incidents and advisories that have been resolved. An incident is a critical service issue, typically involving noticeable user impact. An advisory is a service issue that is typically limited in scope or impact."
                            />
                        </div>
                    </div>
                    <div className="row">
                        <div className="col">
                            <div>
                                <CommandBar items={_commandBarItems} farItems={_commandBarFarItems} style={{ borderTop: "1px solid rgb(225, 225, 225)" }} />
                            </div>
                            <DetailsList
                                items={items}
                                compact={false}
                                columns={columns}
                                selectionMode={SelectionMode.none}
                                getKey={this._getKey}
                                setKey="none"
                                layoutMode={DetailsListLayoutMode.justified}
                                isHeaderVisible={true}
                                onItemInvoked={this._onItemInvoked}
                            />
                            <div className="loadingProgress" style={{ display: isDataLoaded ? 'none' : 'block' }}>
                                <br /><br />
                                <Spinner size={SpinnerSize.large} />
                            </div>

                            <Panel
                                headerText={selectedService}
                                isOpen={isPanelOpen}
                                onDismiss={this._onDismisPanel}
                                type={PanelType.medium}
                                // You MUST provide this prop! Otherwise screen readers will just say "button" with no label.
                                closeButtonAriaLabel="Close"
                            >
                                <DetailsList
                                    compact={true}
                                    selectionMode={SelectionMode.none}
                                    items={selectedFeatureListState}
                                    columns={_featureListcolumns}
                                    layoutMode={DetailsListLayoutMode.fixedColumns}
                                    setKey="none"
                                />

                                <Dropdown
                                    label="Incident"
                                    selectedKey={selectedIncidentItem ? selectedIncidentItem.key : undefined}
                                    // eslint-disable-next-line react/jsx-no-bind
                                    onChange={this._onIncidentDropdownChange}
                                    placeholder="Select an incident"
                                    options={selectedServiceIncidents}
                                    styles={dropdownStyles}
                                    isDisabled={selectedServiceIncidents.length === 0}
                                />
                                <IncidentDetails id={selectedIncidentItem?.key.toString()} />
                            </Panel>

                            <Panel
                                headerText="Incident calender"
                                isOpen={isCalenderPanelOpen}
                                onDismiss={this._onDismisCalenderPanel}
                                type={PanelType.medium}
                                // You MUST provide this prop! Otherwise screen readers will just say "button" with no label.
                                closeButtonAriaLabel="Close"
                            >
                                <IncidentCalendar />        
                            </Panel>
                        </div>
                    </div>
                </div>
            </Fabric>
        );
    }

    public componentDidUpdate(previousProps: any, previousState: IServiceListState) {
        if (previousState.isModalSelection !== this.state.isModalSelection && !this.state.isModalSelection) {
            this._selection.setAllSelected(false);
        }
    }

    componentDidMount() {
        var items: IServiceStatus[] = [];
        var i = 0;
        var accessToken: AuthenticationResult;
        var statistics: IServiceStatistics | undefined = undefined;
        const requiredRoles: string[] = ['ServiceHealthReader', 'Communication.Write.All', 'Admin'];

        acquireAccessToken()
            .then((response) => {
                var tokenClaims: any = response.account?.idTokenClaims;
                const userRoles: any = tokenClaims?.roles;
                var userHasRequiredRole: boolean = userRoles.some((r: string) => requiredRoles.includes(r));

                this.setState({
                    accessGranted: userHasRequiredRole
                });

                accessToken = response;

                if (userHasRequiredRole)
                    fetch('/api/ServiceStatus', { headers: response.idToken === "" ? {} : { 'Authorization': `Bearer ${response.idToken}` } })
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
                            for (const serviceStatus of result) {
                                var fStatus: IFeatureStatus[] = [];
                                var serviceIssues: any = serviceStatus.issues !== undefined ? serviceStatus.issues.filter((i: any) => !i.isResolved) : [];

                                if (serviceIssues !== undefined) {
                                    for (const issue of serviceIssues) {
                                        fStatus.push({
                                            feature: issue.feature,
                                            status: issue.additionalData.status
                                        });
                                    }
                                }

                                items.push({
                                    key: i.toString(),
                                    name: serviceStatus.service,
                                    value: serviceStatus.service,
                                    service: serviceStatus.service,
                                    serviceStatus: serviceStatus.additionalData.status,
                                    issues: serviceIssues,
                                    featureStatus: fStatus
                                });
                                i++;
                            }
                        }).then(() => {
                            fetch('/api/Statistics', { headers: accessToken.idToken === "" ? {} : { 'Authorization': `Bearer ${accessToken.idToken}` } })
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
                                    statistics = result;
                                    /* var sortedEvents = statistics?.eventStatistics.events
                                        .sort((a: IServiceHealthEvent, b: IServiceHealthEvent) => ((a.workload > b.workload ? 1 : -1)))
                                        .sort((a: IServiceHealthEvent, b: IServiceHealthEvent) => ((a.title > b.title ? 1 : -1)));
                                     */
                                }).then(() => {
                                    this.setState({
                                        isDataLoaded: true
                                    });
                                }).catch((err) => {
                                    this.setState({
                                        error: err.message
                                    });
                                });
                        }).then(() => {
                            this._allItems = items;
                            this.setState({
                                items: this._allItems,
                                isDataLoaded: true
                            });
                        });
            }).catch((err) => {
                this.setState({
                    error: err.message
                });
            });
    }

    private _onIncidentDropdownChange = (event: React.FormEvent<HTMLDivElement>, item?: IDropdownOption): void => {
        this.setState({
            selectedIncidentItem: item,
            selectedIncidentItemText: item?.text,
        });
    };

    private _getKey(item: any, index?: number): string {
        return item.key;
    }

    private _onImpactedServicesFilterChange = (checked: boolean): void => {
        const { textFilter } = this.state;
        var filteredItems: IServiceStatus[];

        filteredItems = textFilter ? this._allItems.filter(
            i => i.service.toLowerCase().indexOf(textFilter.toLowerCase()) > -1) : this._allItems;

        if (checked) {
            filteredItems = filteredItems.filter(i => _getServiceHealthState(i.serviceStatus) !== ServiceHealthState.Ok || (i.issues !== undefined && i.issues.length > 0));
        }

        this.setState({
            impactedServicesOnly: checked,
            items: filteredItems,
        });
    }

    private _onChangeSearchText = (ev?: React.ChangeEvent<HTMLInputElement>, newValue?: string): void => {
        const { impactedServicesOnly } = this.state;
        var filteredItems: IServiceStatus[];

        filteredItems = newValue ? this._allItems.filter(
            i => i.service.toLowerCase().indexOf(newValue.toLowerCase()) > -1) : this._allItems;

        if (impactedServicesOnly) {
            filteredItems = filteredItems.filter(i => _getServiceHealthState(i.serviceStatus) !== ServiceHealthState.Ok || (i.issues !== undefined && i.issues.length > 0));
        }

        this.setState({
            textFilter: newValue,
            items: filteredItems,
        });
    };

    private _onItemInvoked(item: any): void {
        alert(`Item invoked: ${item.name}`);
    }

    private _getSelectionDetails(): string {
        const selectionCount = this._selection.getSelectedCount();

        switch (selectionCount) {
            case 0:
                return 'No items selected';
            case 1:
                return '1 item selected: ' + (this._selection.getSelection()[0] as IServiceStatus).name;
            default:
                return `${selectionCount} items selected`;
        }
    }

    private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
        const { columns, items } = this.state;
        const newColumns: IColumn[] = columns.slice();
        const currColumn: IColumn = newColumns.filter(currCol => column.key === currCol.key)[0];
        newColumns.forEach((newCol: IColumn) => {
            if (newCol === currColumn) {
                currColumn.isSortedDescending = !currColumn.isSortedDescending;
                currColumn.isSorted = true;
                this.setState({
                    announcedMessage: `${currColumn.name} is sorted ${currColumn.isSortedDescending ? 'descending' : 'ascending'
                        }`,
                });
            } else {
                newCol.isSorted = false;
                newCol.isSortedDescending = true;
            }
        });
        const newItems = _copyAndSort(items, currColumn.fieldName!, currColumn.isSortedDescending);
        this.setState({
            columns: newColumns,
            items: newItems,
        });
    };

    _onOpenPanel(workload: string) {
        const featureStatus: IFeatureStatus[] | undefined = this._allItems.find(i => i.service.trim().toLowerCase() === workload.trim().toLowerCase())?.featureStatus;
        const issues: any[] | undefined = this._allItems.find(i => i.service.trim().toLowerCase() === workload.trim().toLowerCase())?.issues;
        const incidentsDropdownOptions: IDropdownOption[] = [];

        if (issues !== undefined) {
            for (const issue of issues) {
                if (!issue.isResolved) {
                    incidentsDropdownOptions.push({
                        key: issue.id,
                        text: '[' + issue.id.toUpperCase() + '] ' + issue.title
                    });
                }
            }
        }

        this.setState({
            selectedService: workload,
            selectedFeatureListState: featureStatus ? featureStatus : [],
            selectedServiceIncidents: incidentsDropdownOptions,
            isPanelOpen: true
        });
    }

    private _onDismisPanel = (): void => {
        this.setState({
            selectedService: "",
            selectedFeatureListState: [],
            selectedIncidentItem: undefined,
            selectedIncidentItemText: "",
            isPanelOpen: false
        });
    }

    private _onOpenCalenderPanel = (): void => {
        this.setState({
            isCalenderPanelOpen: true
        });
    }

    private _onDismisCalenderPanel = (): void => {
        this.setState({
            isCalenderPanelOpen: false
        });
    }

    private _getIncidentsGraphicalRepresentation(count: number): string {
        var s = "";
        for (var i = 0; i < count; i++) {
            s += "•";
        }
        return s;
    }
}

function _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
    const key = columnKey as keyof T;
    return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
}

function _getServiceHealthState(status: string): ServiceHealthState {
    var sStatus = status.toLowerCase();
    if (sStatus.includes('degradation') || sStatus.includes('interruption')) {
        return ServiceHealthState.Error;
    } else if (sStatus.includes('investigating') || sStatus.includes('restoring') || sStatus.includes('recovery')) {
        return ServiceHealthState.Warning;
    } else {
        return ServiceHealthState.Ok;
    }
}