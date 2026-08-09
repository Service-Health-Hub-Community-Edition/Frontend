import * as React from 'react';
import { Spinner, SpinnerSize, Text, Separator } from '@fluentui/react';
import { DefaultPalette, DirectionalHint } from '@fluentui/react';
import {
    LineChart, ILineChartPoints, ILineChartDataPoint
    
} from '@fluentui/react-charting';
import { CountAnnotationBar } from '@m365-admin/count-annotation';
import { IServiceStatistics } from './IServiceStatistics';
import { MessageBar, MessageBarType } from '@fluentui/react';
import { acquireAccessToken } from "../../../../auth/AccessTokenHelper";

interface IMonthlyActiveUsersGraphState {
    lineChartData: ILineChartPoints[];
    width: number;
    rawItems: any;
    isDataLoaded: boolean;
    error?: string;
    reportContainer: any;
}

const _colors = [
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

export class MonthlyActiveUsersGraph extends React.Component<{ services?: string[], container?: string }, IMonthlyActiveUsersGraphState> {
    constructor(props: { services?: string[], container?: string }) {
        super(props);

        this.state = {
            isDataLoaded: false,
            width: 800,
            rawItems: [],
            lineChartData: [],
            error: undefined,
            reportContainer: props.container ? document.getElementById(props.container) : undefined
        };

    }

    public render() {
        const {
            isDataLoaded, lineChartData, width, error
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

        return (
                <LineChart
                    data={{ chartTitle: 'Incidents', lineChartData: lineChartData }}
                    width={width}
                    height={200}
                    legendsOverflowText={'more'}
                    legendProps={{ canSelectMultipleLegends: true, allowFocusOnLegends: true }}
                    tickFormat={'%b'}
                />
        );
    }

    componentDidMount() {
        var lineChartData: ILineChartPoints[] = [];
        var rawItems: any;
        
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

        let uri: string = '/api/reports/monthlyActiveUsers';

        if (this.props.services && this.props.services.length > 0)
            uri += '?service=' + this.props.services.join(',');

            acquireAccessToken()
                .then((response) => {
                    fetch(uri,
                        { headers: response.idToken === "" ? {} : { 'Authorization': `Bearer ${response.idToken}` } })
                        .then(response => {
                            if (response.ok) {
                                return response.json();
                            } else {
                                var errorMsg: string = this.state.error === undefined ? "" : this.state.error + "<br/>";
                                this.setState({
                                    error: errorMsg + "Fetching data from /api/reports/monthlyActiveUsers endpoint failed with an error " + response.status + " " + response.statusText
                                });
                                throw Error(response.status + " " + response.statusText);
                            }
                        })
                        .then(result => {
                            if (result && result.length > 0) {
                                rawItems = result;
                                var sortedResult = result.sort((a: any, b: any) => new Date(a.reportDate) > new Date(b.reportDate) ? 1 : -1);
                                var groupedResult = this._groupBy(sortedResult, 'service', null);

                                let i: number = 0;

                                for (const service of Object.keys(groupedResult)) {
                                    var clIndex = service.trim();

                                    var lineChartDataPoints: ILineChartDataPoint[] = [];

                                    for (const record of groupedResult[service]) {

                                        lineChartDataPoints.push({
                                            x: new Date(record.reportDate),
                                            y: record.active ? record.active : 0
                                        });
                                    }

                                    lineChartData.push({
                                        legend: clIndex,
                                        data: lineChartDataPoints,
                                        color: _colors[0][i % _colors[0].length]
                                    });

                                    i++;
                                }
                            }
                        }).
                        then(() => {
                            this.setState({
                                lineChartData: lineChartData,
                                rawItems: rawItems,
                                isDataLoaded: true
                            });
                        });
                }).catch((err) => {
                    var errorMsg: string = this.state.error === undefined ? "" : this.state.error + "<br/>";
                    this.setState({
                        error: errorMsg + "Failed retrieving data from /api/reports/monthlyActiveUsers with an error " + err.message
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