import * as React from 'react';
import { Text, Spinner, SpinnerSize, SelectionMode, DetailsList, DetailsListLayoutMode, IColumn, Image, Link, ILinkStyleProps, ILinkStyles, DefaultPalette, Panel, PanelType, PrimaryButton } from '@fluentui/react';
import { IBreadcrumbItem } from '@fluentui/react';
import { M365Breadcrumb } from '@m365-admin/m365-breadcrumb';
import { DetailPageHeader } from '@m365-admin/detail-page';
import { MessageBar, MessageBarType } from '@fluentui/react';
import { acquireAccessToken } from "../auth/AccessTokenHelper";
import { checkInTeams } from './auth/detectTeams';
import { AccessDenied } from "./AccessDenied";
import { StackedBarChart } from '@fluentui/react-charting';
import { LicenseDetails } from './LicenseDetails';

interface ILicenseStatistics {
    displayName: string;
    skuPartNumber: string;
    date: Date;
    enabled: number;
    consumedUnits: number;
    availableUnits: number;
    warning: number;
    suspended: number;
}

export interface IIncidentReportState {
    selectedSku: string;
    statistics: ILicenseStatistics[];
    columns: IColumn[];
    announcedMessage?: string;
    isPanelOpen: boolean;
    initialized: boolean;
    inTeams: boolean;
    error?: string;
    accessGranted?: boolean;
}

export class LicenseReport extends React.Component<{}, IIncidentReportState> {
    private _breadcrumbItems: IBreadcrumbItem[] = [];

    constructor(props: {}) {
        super(props);

        const pipeFabricStyles = (p: ILinkStyleProps): ILinkStyles => ({
            root: {
                textDecoration: 'none',
                color: p.theme.semanticColors.bodyText,
                fontWeight: '600',
                fontSize: p.theme.fonts.medium.fontSize,
            },
        });

        const columns: IColumn[] = [
            {
                key: 'clName',
                name: 'Name',
                fieldName: 'name',
                minWidth: 120,
                maxWidth: 300,
                isRowHeader: true,
                isResizable: true,
                isSorted: true,
                onColumnClick: this._onColumnClick,
                onRender: (item: ILicenseStatistics) => {
                    return <div>
                        <img src='/images/Office.png' style={{ width: '32px', marginTop: '-4px', float: 'left' }} />
                        <Link onClick={(event) => {
                            event.preventDefault();
                            this._onOpenPanel(item.skuPartNumber);
                        }} styles={pipeFabricStyles}>{item.displayName}</Link >
                    </div>;
                },
                data: 'string',
                isPadded: true,
            },
            {
                key: 'clAvailableLicenses',
                name: 'Available',
                fieldName: 'availableLicenses',
                minWidth: 60,
                maxWidth: 60,
                isResizable: false,
                isCollapsible: false,
                data: 'string',
                isPadded: true,
                onRender: (item: ILicenseStatistics) => {
                    if (item.enabled >= 1000000)
                        return <span>Unlimited</span>
                    else
                        return <span>{item.availableUnits}</span>;
                },
            },
            {
                key: 'clAssignedLicenses',
                name: 'Assigned',
                fieldName: 'assignedLicenses',
                minWidth: 80,
                maxWidth: 80,
                isResizable: false,
                isCollapsible: false,
                data: 'string',
                isPadded: true,
                onRender: (item: ILicenseStatistics) => {
                    return <div>
                        <StackedBarChart data={{
                            chartTitle: "",
                            chartData: [
                                {
                                    legend: "Assigned",
                                    data: item.consumedUnits,
                                    color: DefaultPalette.redDark
                                },
                                {
                                    legend: "Available",
                                    data: item.availableUnits,
                                    color: DefaultPalette.blueLight
                                }
                            ]
                        }} hideLegend={true} hideNumberDisplay={true} hideTooltip={true} hideDenominator={true} />
                    </div>;
                },
            },
            {
                key: 'clConsumed',
                name: 'Consumed',
                fieldName: 'consumed',
                minWidth: 80,
                maxWidth: 80,
                isResizable: true,
                isCollapsible: false,
                data: 'string',
                isPadded: true,
                onRender: (item: ILicenseStatistics) => {
                    if (item.enabled >= 1000000)
                        return <div>
                            <Text variant='small'>{item.consumedUnits}</Text>
                        </div>
                    else
                        return <div>
                            <Text variant='small'>{item.consumedUnits}/{item.enabled+item.warning}</Text>
                        </div>;
                },
            },
            {
                key: 'clWarning',
                name: 'Warning',
                fieldName: 'warning',
                minWidth: 80,
                maxWidth: 80,
                isResizable: true,
                isCollapsible: false,
                data: 'string',
                isPadded: true,
                onRender: (item: ILicenseStatistics) => {
                    return <div>
                        <Text variant='small'>{item.warning}</Text>
                    </div>
                },
            },
            {
                key: 'clSuspended',
                name: 'Suspended',
                fieldName: 'suspended',
                minWidth: 80,
                maxWidth: 80,
                isResizable: true,
                isCollapsible: false,
                data: 'string',
                isPadded: true,
                onRender: (item: ILicenseStatistics) => {
                    return <div>
                        <Text variant='small'>{item.suspended}</Text>
                    </div>
                },
            }
        ];

        this.state = {
            selectedSku: '',
            statistics: [],
            initialized: false,
            isPanelOpen: false,
            announcedMessage: undefined,
            columns: columns,
            inTeams: checkInTeams(),
            error: undefined,
            accessGranted: undefined
        };

        this._breadcrumbItems = this.state.inTeams ? [
            { text: 'Licenses', key: 'licenses', isCurrentItem: true }
        ] : 
            [
                { text: 'Home', key: 'status', href: '/' },
                { text: 'Licenses', key: 'licenses', isCurrentItem: true }
            ];
    }

    public render() {
        const { columns, statistics, error, accessGranted, initialized, isPanelOpen, selectedSku } = this.state;

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

        if (!initialized) {
            return (
                <div className="container" style={{ maxWidth: '1400px' }}>
                    <div className="row">
                        <div className="col">
                            <br />&nbsp;
                        </div>
                    </div>
                    <div className="row">
                        <div className="col">
                            <div className="loadingProgress" style={{ display: error !== undefined ? 'none' : 'block' }}>
                                <br /><br />
                                <Spinner size={SpinnerSize.large} />
                            </div>
                            <div style={{ display: error !== undefined ? 'block' : 'none' }}>
                                <MessageBar
                                    messageBarType={MessageBarType.error}
                                    isMultiline={false}
                                >
                                    Couldn't retrieve data. Error: {error}
                                </MessageBar>
                                <br />
                            </div>
                        </div>
                    </div>
                </div>

            );
        } else {
            var title = statistics ? statistics.find(m => m.skuPartNumber.toLowerCase() === selectedSku.toLowerCase())?.displayName! : "";

            return (
                <div>
                    <div style={{ display: error !== undefined ? 'block' : 'none' }}>
                        <MessageBar
                            messageBarType={MessageBarType.error}
                            isMultiline={false}
                        >
                            Couldn't retrieve data. Error: {error}
                        </MessageBar>
                        <br />
                    </div>

                    <div className="container" style={{ display: error === undefined ? 'block' : 'none' }}>
                        <div className="row">
                            <div className="col">
                                <M365Breadcrumb
                                    items={this._breadcrumbItems}
                                    maxDisplayedItems={10}
                                    ariaLabel="Breadcrumb with items rendered as buttons"
                                    overflowAriaLabel="More links"
                                />
                                <DetailPageHeader
                                    title="Your products"
                                    description="These are the products owned by your organization. Select a product to view license usage and forecast."
                                />
                            </div>
                        </div>
                        
                        {statistics.length > 0 ? (
                            <div className="row">
                                <div className="col" style={{ minWidth: 300 }}>
                                    <DetailsList
                                        items={statistics}
                                        compact={false}
                                        columns={columns}
                                        selectionMode={SelectionMode.none}
                                        setKey="none"
                                        layoutMode={DetailsListLayoutMode.justified}
                                        isHeaderVisible={true}
                                    />
                                    <Panel
                                        headerText={title}
                                        isOpen={isPanelOpen}
                                        onDismiss={this._onDismissPanel}
                                        type={PanelType.large}
                                        // You MUST provide this prop! Otherwise screen readers will just say "button" with no label.
                                        closeButtonAriaLabel="Close"
                                    >
                                        <LicenseDetails skuPart={ selectedSku } skuDisplayName={ title } />
                                    </Panel>
                                </div>
                            </div>) : (
                                <div className="row">
                                    <div className="col" style={{
                                        minWidth: 300,
                                        justifyContent: 'center',
                                        alignItems: 'center',
                                        width: '100%',
                                    }}>
                                        <div className="noLicenseStatisticsAvailable" style={{ textAlign: 'center' }}>
                                            <img src='/images/noData.png' width='300px' /><br /><br />
                                            <Text variant='medium'><b>There are no license statistics available.</b></Text><br />
                                            <Text variant='medium'>Please configure and start license sync job in Azure portal.</Text><br />
                                        </div>
                                    </div>
                                </div>
                            )
                        }
                    </div>
                </div>
            );
        }
    }

    componentDidMount() {
        this._getLicenseStatistics();
    }

   
    private _getLicenseStatistics() {
        
        const requiredRoles: string[] = ['Admin', 'LicenseReader'];
        var statistics: ILicenseStatistics[] = [];
        acquireAccessToken()
            .then((response) => {
                var tokenClaims: any = response.account?.idTokenClaims;
                const userRoles: any = tokenClaims?.roles;
                var userHasRequiredRole: boolean = userRoles.some((r: string) => requiredRoles.includes(r));

                this.setState({
                    accessGranted: userHasRequiredRole
                });

                if (userHasRequiredRole)
                    fetch('/api/LicenseStatistics?select=statistics',
                        { headers: response.idToken === "" ? {} : { 'Authorization': `Bearer ${response.idToken}` } })
                        .then(response => {
                            if (response.ok) {
                                return response.json();
                            } else {
                                var errorMsg: string = this.state.error === undefined ? "" : this.state.error + "<br/>";
                                this.setState({
                                    error: errorMsg + "Fetching data from /api/LicenseStatistics endpoint failed with an error " + response.status + " " + response.statusText
                                });
                                throw Error(response.status + " " + response.statusText);
                            }
                        })
                        .then(result => {
                            var stats = result.statistics;
                            statistics = _copyAndSort(stats, 'displayName', false);
                    }).then(() => {
                        this.setState({
                            statistics: statistics,
                            initialized: true
                        });
                    });
                }).catch((err) => {
                    var errorMsg: string = this.state.error === undefined ? "" : this.state.error + "<br/>";
                    this.setState({
                        error: errorMsg + "Failed retrieving data from /api/LicenseStatistics with an error " + err.message
                    });
                });
    }

    _onOpenPanel(id: string) {
        this.setState({
            selectedSku: id,
            isPanelOpen: true
        });
    }

    private _onDismissPanel = (): void => {
        this.setState({
            selectedSku: "",
            isPanelOpen: false
        });
    }

    private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
        const { columns, statistics } = this.state;
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
        const newItems = _copyAndSort(statistics, currColumn.fieldName!, currColumn.isSortedDescending);
        this.setState({
            columns: newColumns,
            statistics: newItems,
        });
    };
}

function _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
    const key = columnKey as keyof T;
    return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
}