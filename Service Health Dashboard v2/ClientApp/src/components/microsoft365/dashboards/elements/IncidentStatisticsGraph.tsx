import * as React from 'react';
import { Spinner, SpinnerSize, Text, Separator } from '@fluentui/react';
import { DefaultPalette, DirectionalHint } from '@fluentui/react';
import {
    LineChart, ICustomizedCalloutData, ChartHoverCard,
    ILineChartPoints, ILineChartDataPoint, VerticalStackedBarChart, IVSChartDataPoint, IVerticalStackedChartProps
    
} from '@fluentui/react-charting';
import { CountAnnotationBar } from '@m365-admin/count-annotation';
import { IServiceStatistics } from './IServiceStatistics';
import { MessageBar, MessageBarType } from '@fluentui/react';
import { acquireAccessToken } from "../../../../auth/AccessTokenHelper";

interface IIncidentStatisticsGraphState {
    statistics?: IServiceStatistics;
    lineChartData: ILineChartPoints[];
    chartHeight: number;
    width: number;
    rawItems: any;
    isDataLoaded: boolean;
    error?: string;
    reportContainer: any;
}

export class IncidentStatisticsGraph extends React.Component<{ statistics?: IServiceStatistics, container?: string }, IIncidentStatisticsGraphState> {
    constructor(props: { statistics?: IServiceStatistics, container?: string }) {
        super(props);

        this.state = {
            statistics: this.props.statistics,
            isDataLoaded: false,
            width: 800,
            rawItems: [],
            lineChartData: [],
            chartHeight: 0,
            error: undefined,
            reportContainer: props.container ? document.getElementById(props.container) : undefined
        };

    }

    public render() {
        const {
            isDataLoaded, statistics, lineChartData, width, chartHeight, error
        } = this.state;

        if (!isDataLoaded)
        return (
            <div className="loadingProgress" style={{ display: 'block', textAlign: 'center', width: '100%' }}>
                    <br /><br />
                    <Spinner size={SpinnerSize.large} />
                </div>
        );

        if (error)
            return (
                <MessageBar
                    messageBarType={MessageBarType.error}
                    isMultiline={false}
                >
                    Couldn't retrieve data. Error: {error}
                </MessageBar>
            );

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
            <div className="container">
                <div className="row">
                    <div className="col">
                        <Text variant='medium'><b>Past 30 days</b></Text><br />
                    </div>
                </div>
                
                <div className="row">
                    <div className="col" style={{paddingBottom: "24px"}} >
                        <CountAnnotationBar countAnnotationProps={[
                            {
                                annotationText: 'Incidents',
                                count: incidentCount30,
                                annotationColor: 'red'
                            },
                            {
                                annotationText: 'Advisories',
                                count: advisoryCount30,
                                annotationColor: '#0078d4'
                            }
                        ]} />
                    </div>
                </div>
                
                <div className="row">
                    <div className="col">
                        <Text variant='medium'><b>Past 180 days</b></Text><br />
                    </div>
                </div>
                

                <div className="row">
                    <div className="col" style={{ display: 'flex', height: '200px' }} >
                        <LineChart
                            data={{ chartTitle: 'Incidents', lineChartData: lineChartData }}
                            width={width}
                            height={200}
                            legendsOverflowText={'more'}
                            legendProps={{ canSelectMultipleLegends: true, allowFocusOnLegends: true }}
                            tickFormat={'%b'}
                        />
                    </div>
                </div>               
            </div>
        );
    }

    componentDidMount() {
        var statistics: IServiceStatistics | undefined = undefined;
        var lineChartData: ILineChartPoints[] = [];
        var rawItems: any;
        var statsLoaded: boolean = false;
        var chartLoaded: boolean = false;
        
        if (this.props.container) {
            const container: any = this.props.container ? document.getElementById(this.props.container) : undefined;

            this.setState(
                {
                    reportContainer: this.props.container ? document.getElementById(this.props.container) : undefined,
                    width: container && container !== null ? container.getBoundingClientRect().width : 800
                });

            window.addEventListener("resize", (event) => {
                this.setState({
                    width: this.getParentContainerWidth()
                });
            }, false);
        }

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
                fetch('/api/issues?select=startDateTime,classificationDisplayName',
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
                    rawItems = result;
                    var sortedResult = result.sort((a: any, b: any) => new Date(a.startDateTime) > new Date(b.startDateTime) ? 1 : -1);
                    var groupedResult = this._groupBy(sortedResult, 'classificationDisplayName', null);

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
                    error: errorMsg + "Failed retrieving data from /api/issues with an error " + err.message
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

    private getParentContainerWidth(): number {
        var width: number = 800;

        if (this.state.reportContainer && this.state.reportContainer !== null) {
            const rect = this.state.reportContainer?.getBoundingClientRect();
            width = rect.width;
        }

        return width;
    }
}