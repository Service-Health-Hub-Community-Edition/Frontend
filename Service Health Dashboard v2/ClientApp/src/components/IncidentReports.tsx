import * as React from 'react';
import { DefaultButton } from '@fluentui/react';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { DefaultPalette } from '@fluentui/react';
import { IChartProps, IChartDataPoint, LineChart, MultiStackedBarChart, ILineChartPoints, ILineChartDataPoint } from '@fluentui/react-charting';
import { Text } from '@fluentui/react';
import { VerticalStackedBarChart } from '@fluentui/react-charting';
import { IVSChartDataPoint, IVerticalStackedChartProps } from '@fluentui/react-charting';
import { DirectionalHint } from '@fluentui/react';
import { IBreadcrumbItem } from '@fluentui/react';
import { M365Breadcrumb } from '@m365-admin/m365-breadcrumb';
import { DetailPageHeader } from '@m365-admin/detail-page';
import { MessageBar, MessageBarType } from '@fluentui/react';
import { acquireAccessToken } from "../auth/AccessTokenHelper";
import { checkInTeams } from './auth/detectTeams';
import { AccessDenied } from "./AccessDenied";

interface IEventExtraInfo {
    icon: string;
    text: string;
}

interface IEvents {
    name: string;
    date: number;
    allday: boolean;
    extra: IEventExtraInfo;
}

interface IEventDetails {
    id: string;
    title: string;
    service: string;
    start: Date;
    end?: Date;
    duration: string;
}

interface IDate {
    day: number;
    month: number;
    year: number;
}

export interface IIncidentReportState {
    filter: string[];
    items: IEvents[];
    detailItems: IEventDetails[];
    selectedDate: Date;
    initialized: boolean;
    isMessageDetailsDialogOpen: boolean;
    selectedMessage: string;
    lineChartData: ILineChartPoints[];
    serviceIncidentsTotals: IChartProps[];
    rawItems: any;
    inTeams: boolean;
    error?: string;
    accessGranted?: boolean;
}

interface IVerticalChartDataPoint {
    x: string;
    y: number;
    legend: string;
}

export class IncidentReport extends React.Component<{}, IIncidentReportState> {
    private _allItems: IEvents[];
    private _allLineChartData: ILineChartPoints[];
    private _allServiceIncidentsTotals: IChartProps[];
    private _breadcrumbItems: IBreadcrumbItem[] = [];

    private _colors = [
        [
            DefaultPalette.red,
            DefaultPalette.green,
            DefaultPalette.yellow,
            DefaultPalette.blueLight,
            DefaultPalette.black,
            DefaultPalette.magenta,
            DefaultPalette.greenLight,
            DefaultPalette.magenta,
            DefaultPalette.blue,
            DefaultPalette.neutralPrimary,
            DefaultPalette.neutralSecondary,
            DefaultPalette.neutralTertiary,
            DefaultPalette.purpleDark,
            DefaultPalette.purple,
            DefaultPalette.purpleLight,
            DefaultPalette.magentaDark,
            DefaultPalette.magenta,
            DefaultPalette.magentaLight,
            DefaultPalette.redDark,
            DefaultPalette.orange,
            DefaultPalette.orangeLight,
            DefaultPalette.orangeLighter,
            DefaultPalette.yellowLight,
        ]
    ];

    constructor(props: {}) {
        super(props);

        this._allItems = [];
        this._allLineChartData = [];
        this._allServiceIncidentsTotals = [];

        this.state = {
            filter: [],
            items: [],
            detailItems: [],
            selectedDate: new Date(),
            initialized: false,
            isMessageDetailsDialogOpen: false,
            selectedMessage: "",
            lineChartData: [],
            serviceIncidentsTotals: [],
            rawItems: [],
            inTeams: checkInTeams(),
            error: undefined,
            accessGranted: undefined
        };

        this._breadcrumbItems = this.state.inTeams ? [
            { text: 'Status', key: 'status', isCurrentItem: false },
            { text: 'Reports', key: 'reports', isCurrentItem: true }
        ] : 
            [
                { text: 'Home', key: 'status', href: '/' },
                { text: 'Status', key: 'status', isCurrentItem: false },
                { text: 'Reports', key: 'reports', isCurrentItem: true }
            ];
    }

    public render() {
        const { lineChartData, serviceIncidentsTotals, error, accessGranted, initialized } = this.state;

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
        }

        const hideRatio: boolean[] = [true, false];
        const hideDenominator: boolean[] = [true, true];

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
                                style={{ marginBottom: '16px' }}
                                ariaLabel="Breadcrumb with items rendered as buttons"
                                overflowAriaLabel="More links"
                            />
                            <DetailPageHeader
                                title="Event report"
                                description="Number of service events per month. An incident is a critical service issue, typically involving noticeable user impact. An advisory is a service issue that is typically limited in scope or impact."
                            />
                        </div>
                    </div>
                    
                    <div className="row">
                        <div className="col">
                            <MultiStackedBarChart
                                data={serviceIncidentsTotals}
                                hideDenominator={hideDenominator}
                                hideRatio={hideRatio}
                                focusZonePropsForLegendsInHoverCard={{ 'aria-label': 'legends Container' }}
                                legendsOverflowText={'more'}
                            /><br/>&nbsp;
                        </div>
                    </div>
                    <div className="row">
                        <div className="col">
                            <div style={{ borderTop: "1px solid black" }}>

                                <Text variant={'medium'}><b>Events</b></Text>
                                <LineChart
                                    data={{ chartTitle: 'Incidents', lineChartData: lineChartData }}
                                    legendsOverflowText={'more'}
                                    legendProps={{ canSelectMultipleLegends: true, allowFocusOnLegends: true }}
                                /><br/>&nbsp;
                            </div>
                        </div>
                    </div>
                    <div className="row">
                        <div className="col" style={{ minWidth: 300 }}>
                            <div style={{ borderTop: "1px solid black" }}>
                                <Text variant={'medium'}><b>Exchange</b></Text>
                                <VerticalStackedBarChart
                                    data={this._filterRawItems("Exchange Online")}
                                    yAxisTickCount={10}
                                    chartLabel="Monthly breakdown for Exchange Online"
                                    isCalloutForStack={true}
                                    calloutProps={{
                                        directionalHint: DirectionalHint.topCenter,
                                    }}
                                    margins={{ left: 50 }}
                                />
                                <br/><DefaultButton text="View more" onClick={() => this._handleRedirect("/reports/exchange")} /><br/>&nbsp;
                            </div>
                        </div>

                        <div className="col" style={{ minWidth: 300 }}>
                            <div style={{ borderTop: "1px solid black" }}>
                                <Text variant={'medium'}><b>OneDrive for Business</b></Text>
                                <VerticalStackedBarChart
                                    data={this._filterRawItems("OneDrive for Business")}
                                    yAxisTickCount={10}
                                    chartLabel="Monthly breakdown for OneDrive for Business"
                                    isCalloutForStack={true}
                                    calloutProps={{
                                        directionalHint: DirectionalHint.topCenter,
                                    }}
                                    margins={{ left: 50 }}
                                />
                                <br /><DefaultButton text="View more" onClick={() => this._handleRedirect("/reports/onedrive")} /><br />&nbsp;
                            </div>
                        </div>

                        <div className="col" style={{ minWidth: 300 }}>
                            <div style={{ borderTop: "1px solid black" }}>
                                <Text variant={'medium'}><b>SharePoint Online</b></Text>
                                <VerticalStackedBarChart
                                    data={this._filterRawItems("SharePoint Online")}
                                    yAxisTickCount={10}
                                    chartLabel="Monthly breakdown for SharePoint Online"
                                    isCalloutForStack={true}
                                    calloutProps={{
                                        directionalHint: DirectionalHint.topCenter,
                                    }}
                                    margins={{ left: 50 }}
                                />
                                <br /><DefaultButton text="View more" onClick={() => this._handleRedirect("/reports/sharepoint")} /><br />&nbsp;
                            </div>
                        </div>

                        <div className="col" style={{ minWidth: 300 }}>
                            <div style={{ borderTop: "1px solid black" }}>
                                <Text variant={'medium'}><b>Microsoft Teams</b></Text>
                                <VerticalStackedBarChart
                                    data={this._filterRawItems("Microsoft Teams")}
                                    yAxisTickCount={10}
                                    chartLabel="Monthly breakdown for Microsoft Teams"
                                    isCalloutForStack={true}
                                    calloutProps={{
                                        directionalHint: DirectionalHint.topCenter,
                                    }}
                                    margins={{ left: 50 }}
                                />
                                <br /><DefaultButton text="View more" onClick={() => this._handleRedirect("/reports/teams")} /><br />&nbsp;
                            </div>
                        </div>
                    </div>
                    <div className="row">
                        <div className="col" style={{ minWidth: 300 }}>
                            <div style={{ borderTop: "1px solid black" }}>
                                <Text variant={'medium'}><b>Skype for Business</b></Text>
                                <VerticalStackedBarChart
                                    data={this._filterRawItems("Skype for Business")}
                                    yAxisTickCount={10}
                                    chartLabel="Monthly breakdown for Skype for Business"
                                    isCalloutForStack={true}
                                    calloutProps={{
                                        directionalHint: DirectionalHint.topCenter,
                                    }}
                                    margins={{ left: 50 }}
                                />
                                <br /><DefaultButton text="View more" onClick={() => this._handleRedirect("/reports/skype")} /><br />&nbsp;
                            </div>
                        </div>

                        <div className="col" style={{ minWidth: 300 }}>
                            <div style={{ borderTop: "1px solid black" }}>
                                <Text variant={'medium'}><b>Microsoft 365 Suite</b></Text>
                                <VerticalStackedBarChart
                                    data={this._filterRawItems("Microsoft 365 Suite")}
                                    yAxisTickCount={10}
                                    chartLabel="Monthly breakdown for Microsoft 365 Suite"
                                    isCalloutForStack={true}
                                    calloutProps={{
                                        directionalHint: DirectionalHint.topCenter,
                                    }}
                                    margins={{ left: 50 }}
                                />
                                <br /><DefaultButton text="View more" onClick={() => this._handleRedirect("/reports/m365suite")} /><br />&nbsp;
                            </div>
                        </div>
                    </div>
                </div>              
            </div>

        );
    }

    componentDidMount() {
        this._getIncidents();
    }

    private _getIncidents() {
        var lineChartData: ILineChartPoints[] = [];
        var serviceIncidentsTotals: IChartProps[] = [];
        var monthGroups: IVerticalStackedChartProps[] = [];
        var allItems: any = [];
        const requiredRoles: string[] = ['ServiceHealthReader', 'Communication.Write.All', 'Admin'];

        acquireAccessToken()
            .then((response) => {
                var tokenClaims: any = response.account?.idTokenClaims;
                const userRoles: any = tokenClaims?.roles;
                var userHasRequiredRole: boolean = userRoles.some((r: string) => requiredRoles.includes(r));

                this.setState({
                    accessGranted: userHasRequiredRole
                });

                if (userHasRequiredRole)
                    fetch('/api/issues?select=id,title,startDateTime,endDateTime,service,featureGroup,feature,classificationDisplayName',
                        { headers: response.idToken === "" ? {} : { 'Authorization': `Bearer ${response.idToken}` } })
                        .then(response => {
                            if (response.ok) {
                                return response.json();
                            } else {
                                var errorMsg: string = this.state.error === undefined ? "" : this.state.error + "<br/>";
                                this.setState({
                                    error: errorMsg + "Fetching data from /api/issues endpoint failed with an error " + response.status + " " + response.statusText
                                });
                                throw Error(response.status + " " + response.statusText);
                            }
                        })
                    .then(result => {
                        var sIt: IChartDataPoint[] = [];
                        var groupedResult = this._groupBy(result.sort((a: any, b: any) => new Date(a.startDateTime) > new Date(b.startDateTime) ? 1 : -1), 'service', undefined);

                        var i = 0;
                        for (const service of Object.keys(groupedResult)) {
                            sIt.push({
                                legend: service,
                                data: groupedResult[service].length,
                                color: this._colors[0][i % this._colors[0].length],
                                xAxisCalloutData: service,
                                yAxisCalloutData: Math.round(groupedResult[service].length * 100 / result.length) + '%',
                            });

                            i++;
                        }

                        serviceIncidentsTotals.push({
                            chartTitle: 'Service incidents',
                            chartData: sIt.sort((a: IChartDataPoint, b: IChartDataPoint) => a!.data! > b!.data! ? -1 : 1)
                        });

                        var sortedResult = result.sort((a: any, b: any) => new Date(a.startDateTime) > new Date(b.startDateTime) ? 1 : -1);
                        groupedResult = this._groupBy(sortedResult, 'classificationDisplayName', undefined);

                        for (const classification of Object.keys(groupedResult)) {
                            var clIndex = classification.trim().toLowerCase();                  

                            const groups = groupedResult[classification].reduce((groups: any, incident: any) => {
                                var date = incident.startDateTime.split('T')[0];
                                date = date.substring(0, date.lastIndexOf("-"));
                                if (!groups[date]) {
                                    groups[date] = [];
                                }
                                groups[date].push(incident);
                                return groups;
                            }, {});

                            var firstDate = new Date(sortedResult[0].startDateTime);
                            firstDate = new Date(firstDate.getFullYear(), firstDate.getMonth());
                            var lastDate = new Date(sortedResult[sortedResult.length - 1].startDateTime);
                            lastDate = new Date(lastDate.getFullYear(), lastDate.getMonth());
                            var diff = new Date(lastDate.valueOf() - firstDate.valueOf());
                            var year = diff.getFullYear() - 1970 >= 0 ? diff.getFullYear() - 1970 : 0
                            var months = year * 12 + diff.getMonth();

                            var lineChartDataPoints: ILineChartDataPoint[] = [];
                            for (var i = 0; i <= months; i++) {
                                var month = new Date(new Date(firstDate).setMonth(firstDate.getMonth() + i));
                                var monthNum = month.getMonth() + 1;
                                var dateIndex = month.getFullYear() + "-" + (monthNum < 10 ? "0" + monthNum : monthNum);

                                lineChartDataPoints.push({
                                    x: new Date(month),
                                    y: groups[dateIndex] ? groups[dateIndex].length : 0
                                });
                            }

                            var currentDate = new Date();
                            currentDate = new Date(currentDate.getFullYear(), currentDate.getMonth(), 1);
                            var firstDate = new Date(currentDate.setMonth(currentDate.getMonth() - 6));
                            lineChartData.push({
                                legend: classification,
                                data: lineChartDataPoints.filter(d => d.x >= firstDate),
                                color: clIndex === "incident" ? DefaultPalette.red : clIndex === "advisory" ? DefaultPalette.blue : DefaultPalette.green
                            });                    

                            i++;
                        }

                        allItems = result;
                    }).then(() => {
                        this._allLineChartData = lineChartData;
                        this._allServiceIncidentsTotals = serviceIncidentsTotals;
                        this._allItems = allItems;
                        this.setState({
                            lineChartData: this._allLineChartData,
                            serviceIncidentsTotals: this._allServiceIncidentsTotals,
                            rawItems: this._allItems,
                            initialized: true
                        });
                    });
                }).catch((err) => {
                    var errorMsg: string = this.state.error === undefined ? "" : this.state.error + "<br/>";
                    this.setState({
                        error: errorMsg + "Failed retrieving data from /api/issues with an error " + err.message
                    });
                });
    }

    private _handleRedirect(url: string) {
        window.open(url, "_self");
    }

    private _groupBy(xs: any, key: any, key2: any) {
        if (key2 !== undefined) {
            return xs.reduce(function (rv: any, x: any) {
                (rv[x[key][key2]] = rv[x[key][key2]] || []).push(x);
                return rv;
            }, {});
        }
        else {
            return xs.reduce(function (rv: any, x: any) {
                (rv[x[key]] = rv[x[key]] || []).push(x);
                return rv;
            }, {});
        }
    };

    private _filterRawItems(service: string): IVerticalStackedChartProps[] {
        var monthGroups: IVerticalStackedChartProps[] = [];
        if (this.state.rawItems && service !== undefined) {
            const monthlyGroups = this.state.rawItems
                .filter((i: any) => i.service?.trim().toLowerCase() === service?.trim().toLowerCase())
                .reduce((groups: any, incident: any) => {
                    var date = incident.startDateTime.split('T')[0];
                    date = date.substring(0, date.lastIndexOf("-"));
                    if (!groups[date]) {
                        groups[date] = [];
                    }
                    groups[date].push(incident);
                    return groups;
                }, {});

            var monthKeys = Object.keys(monthlyGroups).sort((a, b) => a > b ? 1 : -1);
            for (const month of monthKeys.slice(Math.max(monthKeys.length - 5, 0))) {
                var serviceGroups = this._groupBy(monthlyGroups[month].sort((a: any, b: any) => a.classification > b.classification ? -1 : 1), 'classificationDisplayName', undefined);
                var monthData: IVSChartDataPoint[] = [];
                var i = 0;
                for (const classification of Object.keys(serviceGroups)) {
                    monthData.push({
                        legend: classification.trim(),
                        data: serviceGroups[classification].length,
                        color: classification.trim().toLowerCase() === "incident" ? DefaultPalette.blueDark : DefaultPalette.blue
                    });
                    i++;
                }

                var date = new Intl.DateTimeFormat(window.navigator.language, { month: 'short', year: 'numeric' }).format(new Date(month));

                monthGroups.push({
                    chartData: monthData,
                    xAxisPoint: date
                });
            }

            return monthGroups.length > 0 ? monthGroups : [{ xAxisPoint: "", chartData: [] }];
        }
        else {
            return [{ xAxisPoint: "", chartData: [] }];
        }
    }
}