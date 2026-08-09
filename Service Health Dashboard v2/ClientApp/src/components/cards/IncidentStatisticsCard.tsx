import * as React from 'react';
import { Fabric } from '@fluentui/react';
import { Text } from '@fluentui/react';
import { Link } from '@fluentui/react';
import { Separator } from '@fluentui/react';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { DefaultPalette, DirectionalHint } from '@fluentui/react';
import {
    LineChart,
    ILineChartPoints, ILineChartDataPoint, VerticalStackedBarChart, IVSChartDataPoint, IVerticalStackedChartProps
    
} from '@fluentui/react-charting';
import { InfoCard } from '../InfoCard';
import { StatisticsTile } from '../StatisticsTileComponent';
import { IServiceStatistics } from '../IServiceStatistics';
import { MessageBar, MessageBarType } from '@fluentui/react';
import { acquireAccessToken } from "../../auth/AccessTokenHelper";

interface IIncidentStatisticsCardState {
    icon?: string;
    iconColor: string;
    title?: string;
    href?: string;
    hrefTitle?: string;
    statistics?: IServiceStatistics;
    lineChartData: ILineChartPoints[];
    chartHeight: number;
    rawItems: any;
    isDataLoaded: boolean;
    error?: string;
}

export class IncidentStatisticsCard extends React.Component<{ icon?: string, iconColor?: string, title?: string, href?: string, hrefTitle?: string, statistics?: IServiceStatistics }, IIncidentStatisticsCardState> {
    constructor(props: { icon?: string, iconColor?: string, title?: string, href?: string, hrefTitle?: string, statistics?: IServiceStatistics }) {
        super(props);

        this.state = {
            icon: this.props.icon,
            iconColor: this.props.iconColor ? this.props.iconColor : "#2B2B2B",
            title: this.props.title,
            href: this.props.href,
            hrefTitle: this.props.hrefTitle,
            statistics: this.props.statistics,
            isDataLoaded: false,
            rawItems: [],
            lineChartData: [],
            chartHeight: 0,
            error: undefined
        };

    }

    public render() {
        const {
            icon, iconColor, title, href, hrefTitle, isDataLoaded, statistics, lineChartData, chartHeight, error
        } = this.state;

        var advisoryCount30 = statistics ? (
            statistics.eventStatistics.pastEvents !== undefined ?
                statistics.eventStatistics.pastEvents.advisories30 : 0
        ) : 0;

        var advisoryCount60 = statistics ? (
            statistics.eventStatistics.pastEvents ?
                statistics.eventStatistics.pastEvents.advisories60 : 0
        ) : 0;

        var incidentCount30 = statistics ? (
            statistics.eventStatistics.pastEvents ?
                statistics.eventStatistics.pastEvents.incidents30 : 0
        ) : 0;

        var incidentCount60 = statistics ? (
            statistics.eventStatistics.pastEvents ?
                statistics.eventStatistics.pastEvents.incidents60 : 0
        ) : 0;

        return (
            <Fabric>
            <InfoCard
                    icon={icon}
                    iconColor={iconColor}
                    title={title}
                    href={href}
                    hrefTitle={hrefTitle}
            >
                <div className="loadingProgress" style={{ display: isDataLoaded || error != undefined ? 'none' : 'block' }}>
                    <br /><br />
                    <Spinner size={SpinnerSize.large} />
                </div>

                <div style={{ display: error != undefined ? 'block' : 'none' }}>
                    <MessageBar
                        messageBarType={MessageBarType.error}
                        isMultiline={false}
                    >
                        Couldn't retrieve data. Error: {error}
                    </MessageBar>
                    <br />
                </div>

                <div className="past30Stats" style={{
                        display: isDataLoaded && statistics !== undefined && error === undefined ? 'block' : 'none',
                        width: '100%',
                        height: '330px',
                        overflowX: 'clip',
                        overflowY: 'clip'
                    }}>
                        <div className="past30Header" style={{ textAlign: 'right', width:'100%' }}>
                            <Text variant='xSmall'>Past 30 days</Text><br />
                        </div>

                        <div className="past30Content" style={{ display: 'flex', marginBottom: '-18px'}}>
                            <div className="past30Incidents" style={{ textAlign: 'center', flex: '0 0 50%' }}>
                                <StatisticsTile
                                    value={incidentCount30}
                                    title={incidentCount30 === 1 ? "incident" : "incidents"}
                                    oldValue={incidentCount60} />
                            </div>
                            <div className="past30Incidents" style={{ textAlign: 'center', flex: '1' }}>
                                <StatisticsTile
                                    value={advisoryCount30}
                                    title={advisoryCount30 === 1 ? "advisory" : "advisories"}
                                    oldValue={advisoryCount60} />
                            </div>
                        </div>

                        <Separator />

                        <div className="past180Header" style={{ textAlign: 'right', width: '100%' }}>
                            <Text variant='xSmall'>Past 180 days</Text><br />
                        </div>

                        <div className="past180Content" style={{ display: 'flex', height: '155px' }}>
                            <LineChart
                                data={{ chartTitle: 'Incidents', lineChartData: lineChartData }}
                                width={280}
                                height={chartHeight}
                                legendsOverflowText={'more'}
                                legendProps={{ canSelectMultipleLegends: true, allowFocusOnLegends: true }}
                                tickFormat={'%b'}
                            />                          
                        </div>

                        <div className="incidentStatisticsFooter" style={{ textAlign: 'right', width: '100%' }}>
                            <Link href='/reports' style={{ fontSize: 'x-small' }}>More...</Link>
                        </div>                        
                </div>
                </InfoCard>
            </Fabric>
        );
    }

    componentDidMount() {
        var statistics: IServiceStatistics | undefined = undefined;
        var lineChartData: ILineChartPoints[] = [];
        var rawItems: any;
        var statsLoaded: boolean = false;
        var chartLoaded: boolean = false;

        if (this.state.statistics === undefined) {
            acquireAccessToken()
                .then((response) => {
                    fetch('/api/Statistics', { headers: response.idToken === "" ? {} : { 'Authorization': `Bearer ${response.idToken}` } })
                        .then(response => {
                            if (response.ok) {
                                return response.json();
                            } else {
                                var errorMsg: string = this.state.error === undefined ? "" : this.state.error + "<br/>";
                                this.setState({
                                    error: errorMsg + "Fetching data from /api/Statistics endpoint failed with an error " + response.status + " " + response.statusText
                                });
                                throw Error(response.status + " " + response.statusText);
                            }
                        })
                        .then(result => {
                            statistics = result;
                            statsLoaded = true;
                        }).then(() => {
                            this.setState({
                                statistics: statistics,
                                isDataLoaded: statsLoaded && chartLoaded
                            });
                        }).catch((err) => {
                            this.setState({
                                error: err.message
                            });
                        });
                }).catch((err) => {
                    var errorMsg: string = this.state.error === undefined ? "" : this.state.error + "<br/>";
                    this.setState({
                        error: errorMsg + "Failed retrieving data from /api/Statistics with an error " + err.message
                    });
                });
        }
        else {
            statsLoaded = true;
            this.setState({
                isDataLoaded: statsLoaded && chartLoaded
            });
        }


        acquireAccessToken()
            .then((response) => {
                fetch('/api/Issues?select=startDateTime,classification',
                    { headers: response.idToken === "" ? {} : { 'Authorization': `Bearer ${response.idToken}` } })
                    .then(response => {
                        if (response.ok) {
                            return response.json();
                        } else {
                            var errorMsg: string = this.state.error === undefined ? "" : this.state.error + "<br/>";
                            this.setState({
                                error: errorMsg + "Fetching data from /api/Issues endpoint failed with an error " + response.status + " " + response.statusText
                            });
                            throw Error(response.status + " " + response.statusText);
                        }
                    })
                .then(result => {
                    rawItems = result;
                    var sortedResult = result.sort((a: any, b: any) => new Date(a.startDateTime) > new Date(b.startDateTime) ? 1 : -1);
                    var groupedResult = this._groupBy(sortedResult, 'additionalData', 'classification');

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
                            var date = new Intl.DateTimeFormat(window.navigator.language, { month: 'short', year: 'numeric' }).format(new Date(month));
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
                    chartLoaded = true;
                }).
                then(() => {
                    this.setState({
                        lineChartData: lineChartData,
                        rawItems: rawItems,
                        chartHeight: 155,
                        isDataLoaded: statsLoaded && chartLoaded
                    });
                });
            }).catch((err) => {
                var errorMsg: string = this.state.error === undefined ? "" : this.state.error + "<br/>";
                this.setState({
                    error: errorMsg + "Failed retrieving data from /api/Issues with an error " + err.message
                });
            });
    }

    private _groupBy(xs: any, key: any, key2: any) {
        if (key2 != undefined) {
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
}