import * as React from 'react';
import { DefaultPalette } from '@fluentui/react';
import { IChartProps, LineChart, StackedBarChart, ILineChartPoints, ILineChartDataPoint } from '@uifabric/charting';
import { Text, Spinner, SpinnerSize } from '@fluentui/react';
import { MessageBar, MessageBarType } from '@fluentui/react';
import { acquireAccessToken } from "../auth/AccessTokenHelper";
import { checkInTeams } from './auth/detectTeams';
import { AccessDenied } from "./AccessDenied";

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

interface ILicenseStatisticsExt {
    displayName: string;
    skuPartNumber: string;
    timestamp: Date;
    enabled: number;
    consumedUnits: number;
    suspended: number;
    warning: number;
}


interface ILicenseForecast {
    skuPartNumber: string;
    month: Date;
    newHires: number;
    leavers: number;
    balance: number;
}

export interface ILicenseDetailsState {
    skuDisplayName: string;
    skuPart: string;
    statistics: ILicenseStatistics[];
    forecast: ILicenseForecast[];
    history: ILicenseStatisticsExt[];
    lastSync: Date;
    initialized: boolean;
    statisticsChartData?: IChartProps;
    forecastChartData: ILineChartPoints[];
    historyChartData: ILineChartPoints[];
    inTeams: boolean;
    error?: string;
    accessGranted?: boolean;
}


export class LicenseDetails extends React.Component<{skuPart: string, skuDisplayName: string}, ILicenseDetailsState> {

    constructor(props: {skuPart: string, skuDisplayName: string}) {
        super(props);

        this.state = {
            skuPart: this.props.skuPart,
            skuDisplayName: this.props.skuDisplayName,
            statistics: [],
            forecast: [],
            history: [],
            lastSync: new Date(),
            initialized: false,
            statisticsChartData: undefined,
            forecastChartData: [],
            historyChartData: [],
            inTeams: checkInTeams(),
            error: undefined,
            accessGranted: undefined
        };
    }

    public render() {
        const { statistics, statisticsChartData, forecastChartData, error, accessGranted, initialized, historyChartData } = this.state;

        var consumedUnits: number = 0;
        var totalUnits: number = 0;

        if (statistics !== undefined && statistics.length > 0) {
            consumedUnits = statistics[0].consumedUnits;
            totalUnits = statistics[0].enabled + statistics[0].warning;
        }

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

                    <div className="container" style={{ display: error === undefined ? 'block' : 'none', marginTop: '16px' }}>
                        {statisticsChartData !== undefined ? (
                                <div className="row">
                                    <div className="col" style={{ minWidth: 300 }}>
                                        <div style={{ borderTop: "1px solid black" }}>
                                            <Text variant={'medium'}><b>License usage</b></Text><br /><br />
                                            <Text variant={'large'}>{consumedUnits} / {totalUnits} assigned</Text>
                                            <StackedBarChart data={statisticsChartData} ignoreFixStyle={true} />
                                            <br />&nbsp;
                                        </div>
                                    </div>

                                    <div className="col" style={{ minWidth: 300 }}>
                                        <div style={{ borderTop: "1px solid black" }}>
                                            <Text variant={'medium'}><b>History</b></Text>
                                            <LineChart
                                                data={{ chartTitle: 'License usage', lineChartData: historyChartData ? historyChartData : [] }}
                                                legendsOverflowText={'more'}
                                                legendProps={{ canSelectMultipleLegends: true, allowFocusOnLegends: true }}
                                            />
                                            <br />&nbsp;
                                        </div>
                                </div>
                                </div>
                        ) : (
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
                        {statisticsChartData !== undefined ? (
                            <div className="row">
                                <div className="col style={{ minWidth: 300 }}">
                                    <div style={{ borderTop: "1px solid black" }}>
                                        <Text variant={'medium'}><b>Forecast</b></Text>
                                        {forecastChartData.length > 0 ? (
                                            <LineChart
                                                data={{ chartTitle: 'License forecast', lineChartData: forecastChartData ? forecastChartData : [] }}
                                                legendsOverflowText={'more'}
                                                legendProps={{ canSelectMultipleLegends: true, allowFocusOnLegends: true }}
                                                allowMultipleShapesForPoints={true}
                                            />) : (
                                            <div className="noHistoryDataAvailable" style={{ textAlign: 'center' }}>
                                                <img src='/images/noData.png' width='200 px' /><br /><br />
                                                <Text variant='medium'><b>There is no forecast available for selected SKU.</b></Text><br />
                                                <Text variant='smallPlus'>Forecast is not available for SKUs with unlimited licenses. Please select another SKU and make sure the HR statistics are imported.</Text><br />
                                            </div>
                                        )
                                        } <br />&nbsp;
                                    </div>
                                </div>                           
                            </div>) : (<div />)}
                    </div>
                </div>
            );
        }
    }

    componentDidMount() {
        this._getLicenseStatistics();
    }

    private setCharts = (forecast: ILicenseForecast[], history: ILicenseStatisticsExt[], statisticsChartData: IChartProps | undefined): void => {
        var sortedResult = forecast?.sort((a: any, b: any) => new Date(a.month) > new Date(b.month) ? 1 : -1);

        var balancePoints: ILineChartDataPoint[] = [];
        var newHiresPoints: ILineChartDataPoint[] = [];
        var leaversPoints: ILineChartDataPoint[] = [];
        var lineChartData: ILineChartPoints[] = [];

        if (sortedResult !== undefined) {
            for (const forecastItem of sortedResult) {
                balancePoints.push({
                    x: new Date(forecastItem.month),
                    y: forecastItem.balance
                });

                newHiresPoints.push({
                    x: new Date(forecastItem.month),
                    y: forecastItem.newHires
                });

                leaversPoints.push({
                    x: new Date(forecastItem.month),
                    y: forecastItem.leavers
                });
            }

            if (sortedResult.length > 0) {
                lineChartData.push({
                    legend: 'New employees',
                    data: newHiresPoints,
                    color: DefaultPalette.green
                });

                lineChartData.push({
                    legend: 'Leaving company',
                    data: leaversPoints,
                    color: DefaultPalette.redDark
                });

                lineChartData.push({
                    legend: 'Available units',
                    data: balancePoints,
                    color: DefaultPalette.blue
                });
            }
        }

        var sortedHistoryResult = history?.sort((a: any, b: any) => new Date(a.timestamp) > new Date(b.timestamp) ? 1 : -1);

        var consumedUnits: ILineChartDataPoint[] = [];
        var availableUnits: ILineChartDataPoint[] = [];
        var warnings: ILineChartDataPoint[] = [];
        var suspended: ILineChartDataPoint[] = [];
        var historyLineChartData: ILineChartPoints[] = [];

        if (sortedHistoryResult !== undefined) {
            for (const historyItem of sortedHistoryResult) {
                consumedUnits.push({
                    x: new Date(historyItem.timestamp),
                    y: historyItem.consumedUnits
                });

                availableUnits.push({
                    x: new Date(historyItem.timestamp),
                    y: historyItem.enabled - historyItem.consumedUnits
                });

                warnings.push({
                    x: new Date(historyItem.timestamp),
                    y: historyItem.warning
                });

                suspended.push({
                    x: new Date(historyItem.timestamp),
                    y: historyItem.suspended
                });
            }

            historyLineChartData.push({
                legend: 'Consumed units',
                data: consumedUnits,
                color: DefaultPalette.black
            });

            historyLineChartData.push({
                legend: 'Available units',
                data: availableUnits,
                color: DefaultPalette.green
            });

            historyLineChartData.push({
                legend: 'Warning',
                data: warnings,
                color: DefaultPalette.yellowDark
            });

            historyLineChartData.push({
                legend: 'Suspended',
                data: suspended,
                color: DefaultPalette.orange
            });
        }

        this.setState({
            forecastChartData: lineChartData,
            historyChartData: historyLineChartData,
            statisticsChartData: statisticsChartData
        });
    };

    private _getLicenseStatistics() {
        var statisticsChartData: IChartProps | undefined = undefined;
        var forecast: ILicenseForecast[] = [];
        var history: ILicenseStatisticsExt[] = [];
        var statistics: ILicenseStatistics[] = [];
        const requiredRoles: string[] = ['Admin', 'LicenseReader'];

        acquireAccessToken()
            .then((response) => {
                var tokenClaims: any = response.account?.idTokenClaims;
                const userRoles: any = tokenClaims?.roles;
                var userHasRequiredRole: boolean = userRoles.some((r: string) => requiredRoles.includes(r));

                this.setState({
                    accessGranted: userHasRequiredRole
                });

                if (userHasRequiredRole)
                    fetch('/api/LicenseStatistics?sku=' + this.state.skuPart,
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
                            forecast = result.forecast;
                            history = result.history;
                            statistics = result.statistics;

                            statisticsChartData = {
                                chartTitle: "",
                                chartData: [
                                    {
                                        legend: "Assigned",
                                        data: statistics !== undefined && statistics.length > 0 ? statistics[0].consumedUnits : 0,
                                        color: DefaultPalette.redDark,
                                        xAxisCalloutData: 'Assigned',
                                        yAxisCalloutData: (statistics !== undefined && statistics.length > 0 ? statistics[0].consumedUnits : 0).toString()
                                    },
                                    {
                                        legend: "Available",
                                        data: statistics !== undefined && statistics.length > 0 ? statistics[0].availableUnits : 0,
                                        color: DefaultPalette.blueLight,
                                        xAxisCalloutData: 'Available',
                                        yAxisCalloutData: (statistics !== undefined && statistics.length > 0 ? statistics[0].availableUnits : 0).toString()
                                    },
                                    {
                                        legend: "Warning",
                                        data: statistics !== undefined && statistics.length > 0 ? statistics[0].warning : 0,
                                        color: DefaultPalette.orange,
                                        xAxisCalloutData: 'Warning',
                                        yAxisCalloutData: (statistics !== undefined && statistics.length > 0 ? statistics[0].warning : 0).toString()
                                    }
                                ]
                            };
                        }).then(() => {
                            this.setCharts(forecast, history, statisticsChartData);
                            this.setState({
                                forecast: forecast,
                                history: history,
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
}