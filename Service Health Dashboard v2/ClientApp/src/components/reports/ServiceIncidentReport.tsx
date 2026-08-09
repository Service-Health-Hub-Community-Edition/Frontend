import * as React from 'react';
import { DefaultPalette } from '@fluentui/react';
import { IChartDataPoint, ILineChartPoints, ILineChartDataPoint, LineChart, AreaChart, VerticalBarChart, IVerticalBarChartDataPoint } from '@fluentui/react-charting';
import { Text, Link, ILinkStyleProps, ILinkStyles } from '@fluentui/react';
import { IColumn, DetailsList, SelectionMode, DetailsListLayoutMode } from '@fluentui/react';
import { DonutChart } from '@fluentui/react-charting';
import { IBreadcrumbItem, CommandBar, ICommandBarItemProps } from '@fluentui/react';
import { M365Breadcrumb } from '@m365-admin/m365-breadcrumb';
import { DetailPageHeader } from '@m365-admin/detail-page';
import { Panel, PanelType } from '@fluentui/react';
import { IncidentDetails } from '../IncidentDetails';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { MessageBar, MessageBarType } from '@fluentui/react';
import { acquireAccessToken, accessControl } from "../../auth/AccessTokenHelper";
import { checkInTeams } from '../auth/detectTeams';
import { AccessDenied } from "../AccessDenied";

interface IEventDetails {
    id: string;
    title: string;
    service: string;
    start: Date;
    end?: Date;
    duration: string;
    featureGroup: string;
    feature: string;
    classification: string;
}

enum IIncidentFilter {
    All = 0,
    Incidents = 1,
    Advisories = 2
}

export interface IServiceIncidentReportState {
    service?: string;
    serviceDisplayName?: string;
    lineChartData: ILineChartPoints[];
    featureIncidentsBreakdown: IChartDataPoint[];
    featureGroupsIncidentsBreakdown: IVerticalBarChartDataPoint[];
    incidentList: IEventDetails[];
    incidentFilter: IIncidentFilter;
    isDetailsPanelOpen: boolean;
    selectedIncident: string;
    isDataLoaded: boolean;
    inTeams: boolean;
    error?: string;
    accessGranted?: boolean;
}

export class ServiceIncidentReport extends React.Component<{service: string, serviceDisplayName: string}, IServiceIncidentReportState> {
    private _allLineChartData: ILineChartPoints[] = [];
    private _allFeatureIncidentsBreakdown: IChartDataPoint[] = [];
    private _allFeatureGroupsIncidentsBreakdown: IVerticalBarChartDataPoint[] = [];
    private _allIncidents: IEventDetails[] = [];

    private _colors = [
        [
            DefaultPalette.blue,
            DefaultPalette.green,
            DefaultPalette.yellow,
            DefaultPalette.blueLight,
            DefaultPalette.black,
            DefaultPalette.magenta,
            DefaultPalette.greenLight,
            DefaultPalette.magenta,
            DefaultPalette.red,
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
        ],
        [
            DefaultPalette.orange,
            DefaultPalette.yellow,
            DefaultPalette.blue,
            DefaultPalette.blueMid,
            DefaultPalette.blueLight,
            DefaultPalette.green,
            DefaultPalette.greenLight,
        ],
    ];

    private _breadcrumbItems: IBreadcrumbItem[] = [];

    constructor(props: { service: string, serviceDisplayName: string }) {
        super(props);

        this.state = {
            service: props.service,
            serviceDisplayName: props.serviceDisplayName,
            lineChartData: [],
            featureIncidentsBreakdown: [],
            featureGroupsIncidentsBreakdown: [{ x: "", y: 0 }],
            incidentList: [],
            incidentFilter: IIncidentFilter.All,
            isDetailsPanelOpen: false,
            selectedIncident: "",
            isDataLoaded: false,
            inTeams: checkInTeams(),
            error: undefined,
            accessGranted: undefined
        };

        this._breadcrumbItems = this.state.inTeams ? [
            { text: 'Status', key: 'status', isCurrentItem: false },
            { text: 'Reports', key: 'reports', href: "/reports" },
            { text: this.state.serviceDisplayName ? this.state.serviceDisplayName : "Service", key: "service", isCurrentItem: true }
        ] : 
            [
                { text: 'Home', key: 'home', href: '/' },
                { text: 'Status', key: 'status', isCurrentItem: false },
                { text: 'Reports', key: 'reports', href: "/reports" },
                { text: this.state.serviceDisplayName ? this.state.serviceDisplayName : "Service", key: "service", isCurrentItem: true }
            ];
    }

    public render() {

        const {
            accessGranted
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

        const { lineChartData, featureIncidentsBreakdown, serviceDisplayName, featureGroupsIncidentsBreakdown, incidentList, incidentFilter,
            isDetailsPanelOpen, selectedIncident, isDataLoaded, error } = this.state;

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
                key: 'clId',
                name: 'Id',
                fieldName: 'id',
                minWidth: 80,
                maxWidth: 80,
                isRowHeader: true,
                isResizable: false,
                data: 'string',
                isPadded: true,
            },
            {
                key: 'clTitle',
                name: 'Title',
                fieldName: 'title',
                minWidth: 180,
                maxWidth: 400,
                isResizable: false,
                isCollapsible: false,
                isMultiline: true,
                data: 'string',
                onRender: (item: IEventDetails) => {
                    return <Link onClick={(event) => {
                        event.preventDefault();
                        this._onOpenDetailsPanel(item.id);
                    }} styles={pipeFabricStyles}>{item.title}</Link >;
                },
                isPadded: true,
            },
            {
                key: 'clStartTime',
                name: 'Start',
                fieldName: 'start',
                minWidth: 60,
                maxWidth: 80,
                isResizable: true,
                isCollapsible: true,
                data: 'string',
                isPadded: true,
                onRender: (item: IEventDetails) => {
                    return <span>{item.start ? item.start.toLocaleDateString(): ""}</span>;
                },
            },
            {
                key: 'clEndTime',
                name: 'End',
                fieldName: 'end',
                minWidth: 60,
                maxWidth: 80,
                isResizable: true,
                isCollapsible: true,
                data: 'string',
                isPadded: true,
                onRender: (item: IEventDetails) => {
                    return <span>{item.end ? item.end.toLocaleDateString() : ""}</span>;
                },
            },
            {
                key: 'clDuration',
                name: 'Duration',
                fieldName: 'duration',
                minWidth: 80,
                maxWidth: 100,
                isResizable: true,
                isCollapsible: true,
                data: 'string',
                isPadded: true,
            },
            {
                key: 'clFeature',
                name: 'Feature',
                fieldName: 'feature',
                minWidth: 80,
                maxWidth: 200,
                isResizable: true,
                isCollapsible: true,
                data: 'string',
                isPadded: true,
            },
            {
                key: 'clSeverity',
                name: 'Severity',
                fieldName: 'severity',
                minWidth: 60,
                maxWidth: 80,
                isResizable: true,
                isCollapsible: true,
                data: 'string',
                isPadded: true,
            },
            {
                key: 'clClassification',
                name: 'Classification',
                fieldName: 'classification',
                minWidth: 80,
                maxWidth: 100,
                isResizable: true,
                isCollapsible: true,
                data: 'string',
                isPadded: true,
            }
        ];

        const _commandBarItems: ICommandBarItemProps[] = [];

        const _commandBarFarItems: ICommandBarItemProps[] = [
            {
                key: 'itemsCount',
                onRender: () => {
                    return <div style={{ paddingTop: 12, paddingRight: 8 }}><Text><b>{incidentList.length} item{incidentList.length !== 1 ? "s" : ""}</b></Text></div>;
                },
            },
            {
                key: 'incidentFilter',
                text: IIncidentFilter[incidentFilter],
                iconProps: { iconName: 'List' },
                subMenuProps: {
                    items: [
                        {
                            key: 'all',
                            text: 'All',
                            checked: incidentFilter === IIncidentFilter.All,
                            onClick: () => this._onIncidentFilterChange(IIncidentFilter.All)
                        },
                        {
                            key: 'advisories',
                            text: 'Advisories',
                            checked: incidentFilter === IIncidentFilter.Advisories,
                            onClick: () => this._onIncidentFilterChange(IIncidentFilter.Advisories)
                        },
                        {
                            key: 'incidents',
                            text: 'Incidents',
                            checked: incidentFilter === IIncidentFilter.Incidents,
                            onClick: () => this._onIncidentFilterChange(IIncidentFilter.Incidents)
                        }
                    ]
                },
            }
        ];

        return (
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
                            title={serviceDisplayName ? serviceDisplayName : ""}
                            description={"Number of " + (serviceDisplayName ? serviceDisplayName : "") + " events per month. An incident is a critical service issue, typically involving noticeable user impact. An advisory is a service issue that is typically limited in scope or impact."}
                        />
                    </div>
                </div>

                <div className="loadingProgress" style={{ display: isDataLoaded || error !== undefined ? 'none' : 'block' }}>
                    <br />
                    <Spinner size={SpinnerSize.large} />
                    <br />&nbsp;
                </div>

                <div className="row" style={{ display: error !== undefined ? 'block' : 'none' }}>
                    <MessageBar
                        messageBarType={MessageBarType.error}
                        isMultiline={false}
                    >
                        Couldn't retrieve data. Error: {error}
                    </MessageBar>
                    <br />
                </div>

                <div className="row" style={{ display: error === undefined ? '' : 'none' }}>
                    <div className="col" style={{ minWidth: 300 }}>
                        <div style={{ borderTop: "1px solid black" }}>

                            <Text variant={'medium'}><b>Events</b></Text>
                            <LineChart
                                data={{ chartTitle: 'Incidents', lineChartData: lineChartData ? lineChartData : [] }}
                                legendsOverflowText={'more'}
                                legendProps={{ canSelectMultipleLegends: true, allowFocusOnLegends: true }}
                            /><br/>&nbsp;
                        </div>
                    </div>

                    <div className="col" style={{ minWidth: 300 }}>
                        <div style={{ borderTop: "1px solid black" }}>
                            <Text variant={'medium'}><b>Breakdown by feature groups</b></Text>
                            <VerticalBarChart data={featureGroupsIncidentsBreakdown} width={300} height={200} chartLabel={'Feature groups breakdown'} legendProps={{
                                allowFocusOnLegends: true,
                            }}
                                rotateXAxisLables={true} wrapXAxisLables={true} />
                        </div>

                    </div>

                    <div className="col" style={{ minWidth: 300 }}>
                        <div style={{ borderTop: "1px solid black" }}>
                        <Text variant={'medium'}><b>Breakdown by feature</b></Text>
                <DonutChart
                                data={{ chartTitle: 'Feature breakdown', chartData: featureIncidentsBreakdown }}
                                innerRadius={75}
                                legendProps={{
                                    allowFocusOnLegends: true,
                                }}
                                width={300} height={300}
                                valueInsideDonut={incidentList.length}
                            />
                            </div>
                    </div>
                </div>
                <div className="row" style={{ display: error === undefined ? '' : 'none' }}>
                    <div className="col">
                        <br/>
                        <div>
                            <CommandBar items={_commandBarItems} farItems={_commandBarFarItems} style={{ borderTop: "1px solid rgb(225, 225, 225)" }}/>
                        </div>

                        <DetailsList
                            items={incidentList}
                            compact={false}
                            columns={columns}
                            selectionMode={SelectionMode.none}
                            layoutMode={DetailsListLayoutMode.justified}
                            isHeaderVisible={true}
                        />

                        <Panel
                            headerText={incidentList.find(m => m.id.toLowerCase() === selectedIncident.toLowerCase())?.title!}
                            isOpen={isDetailsPanelOpen}
                            onDismiss={this._onDismisDetailsPanel}
                            type={PanelType.medium}
                            // You MUST provide this prop! Otherwise screen readers will just say "button" with no label.
                            closeButtonAriaLabel="Close"
                        >
                            <IncidentDetails id={selectedIncident} />
                        </Panel>
                    </div>
                </div>
            </div>
        );
    }

    componentDidMount() {
        this._getIncidents();
    }

    private _getIncidents() {
        if (this.state.service !== undefined) {
            var lineChartData: ILineChartPoints[] = [];
            var featureGroupsIncidentsBreakdown: IVerticalBarChartDataPoint[] = [];
            var featureIncidentsBreakdown: IChartDataPoint[] = [];
            var incidents: IEventDetails[] = [];
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
                        fetch("/api/issues?select=id,title,startDateTime,endDateTime,service,featureGroup,feature,classificationDisplayName&filter=affectedServices/any(service:service eq '" + this.state.service + "')",
                            { headers: response.idToken === "" ? {} : { 'Authorization': `Bearer ${response.idToken}` } })
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
                        var sortedResult = result.sort((a: any, b: any) => new Date(a.startDateTime) > new Date(b.startDateTime) ? 1 : -1);
                        var groupedResult = this._groupBy(sortedResult, 'classificationDisplayName', undefined);

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

                            var featureGroups = this._groupBy(sortedResult.filter((s: any) => s.featureGroup !== undefined && s.featureGroup.trim() !== ""), 'featureGroup', undefined);

                            for (const fGroup of Object.keys(featureGroups)) {
                                featureGroupsIncidentsBreakdown.push({
                                    legend: fGroup,
                                    x: fGroup,
                                    y: featureGroups[fGroup].length,
                                    xAxisCalloutData: fGroup,
                                    yAxisCalloutData: featureGroups[fGroup].length
                                });
                            }

                            featureGroupsIncidentsBreakdown = featureGroupsIncidentsBreakdown.sort((a: IVerticalBarChartDataPoint, b: IVerticalBarChartDataPoint) => a.x > b.x ? 1 : -1);
                            var featureGroupCount = 0;

                            for (const featureGroupStat of featureGroupsIncidentsBreakdown) {
                                featureGroupStat.color = this._colors[1][featureGroupCount % this._colors[1].length];
                                featureGroupCount++;
                            }

                            var featureBreakdown = this._groupBy(sortedResult, 'feature', undefined);

                            var featureCount = 0;
                            for (const feature of Object.keys(featureBreakdown)) {
                                featureIncidentsBreakdown.push({
                                    legend: feature,
                                    data: featureBreakdown[feature].length,
                                    color: this._colors[0][featureCount % this._colors[0].length]
                                });
                                featureCount++;
                            }

                            sortedResult = sortedResult.sort((a: any, b: any) => new Date(a.startDateTime) > new Date(b.startDateTime) ? -1 : 1);
                            for (const incident of sortedResult) {
                                var start: Date = new Date(incident.startDateTime);
                                var end: Date | undefined = incident.endDateTime ? new Date(incident.endDateTime) : undefined;
                                var durationNum = end ? end.valueOf() - start.valueOf() : -1;
                                var durationDays = Math.trunc(durationNum / (1000 * 60 * 60 * 24));
                                var duration: string = end ? (durationDays > 0 ? durationDays+"d " : "") + new Intl.DateTimeFormat(
                                    window.navigator.language, { hour: "2-digit", minute: "2-digit", second: "2-digit", hour12: false }).format(durationNum) : "ongoing";

                                incidents.push({
                                    id: incident.id,
                                    title: incident.title,
                                    service: incident.service,
                                    start: start,
                                    end: end,
                                    duration: duration,
                                    featureGroup: incident.featureGroup,
                                    feature: incident.feature,
                                    classification: incident.classificationDisplayName
                                });
                            }
                        }).then(() => {
                            this._allLineChartData = lineChartData;
                            this._allFeatureIncidentsBreakdown = featureIncidentsBreakdown;
                            this._allFeatureGroupsIncidentsBreakdown = featureGroupsIncidentsBreakdown;
                            this._allIncidents = incidents;
                            this.setState({
                                lineChartData: this._allLineChartData,
                                featureIncidentsBreakdown: this._allFeatureIncidentsBreakdown,
                                featureGroupsIncidentsBreakdown: this._allFeatureGroupsIncidentsBreakdown,
                                incidentList: this._allIncidents,
                                isDataLoaded: true
                            });
                        });
                    }).catch((err) => {
                        this.setState({
                            error: err.message
                        });
                    });
        }
    }

    private _onIncidentFilterChange = (filter: IIncidentFilter): void => {
        var filteredItems: IEventDetails[];

        filteredItems = filter === IIncidentFilter.Advisories ?
            this._allIncidents.filter(i => i.classification.trim().toLowerCase() === "advisory") :
            filter === IIncidentFilter.Incidents ?
                this._allIncidents.filter(i => i.classification.trim().toLowerCase() === "incident") : this._allIncidents;
            

        this.setState({
            incidentList: filteredItems,
            incidentFilter: filter
        });
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

    private _randomColor(index: number): string {
        return this._colors[index][Math.floor(Math.random() * this._colors[index].length)];
    }

    _onOpenDetailsPanel(id: string) {
        this.setState({
            selectedIncident: id,
            isDetailsPanelOpen: true
        });
    }

    private _onDismisDetailsPanel = (): void => {
        this.setState({
            selectedIncident: "",
            isDetailsPanelOpen: false
        });
    }
}