import * as React from 'react';
import {
    DefaultPalette, Text, FontIcon, Spinner, SpinnerSize,
    MessageBar, MessageBarType, mergeStyleSets, mergeStyles
} from '@fluentui/react';
import { DonutChart, IChartDataPoint, } from '@fluentui/react-charting';
import { acquireAccessToken } from "../../../../auth/AccessTokenHelper";
import { checkInTeams } from '../../../auth/detectTeams';
import { AccessDenied } from "../../../AccessDenied";

export interface IServiceStatisticsState {
    items: IChartDataPoint[];
    isDataLoaded: boolean;
    inTeams: boolean;
    error?: string;
    accessGranted?: boolean;
}

export enum HealthReportType {
    active,
    history
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

const iconClass = mergeStyles({
    fontSize: 110,
    height: 110,
    width: 110
});

const iconClasses = mergeStyleSets({
    Ok: [{ color: 'green' }, iconClass],
    Degradation: [{ color: 'orange' }, iconClass]
});

export class ServiceStatisticsGraph extends React.Component<{}, IServiceStatisticsState> {

    constructor(props: {}) {
        super(props);

        this.state = {
            items: [],
            isDataLoaded: false,
            inTeams: checkInTeams(),
            error: undefined,
            accessGranted: undefined
        };
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

        const { items, isDataLoaded, error } = this.state;

        if (!isDataLoaded)
            return (<div className="loadingProgress">
                <br />
                <Spinner size={SpinnerSize.large} />
                <br />&nbsp;
            </div>);

        if (error)
            return (
                <MessageBar
                    messageBarType={MessageBarType.error}
                    isMultiline={false}
                >
                    Couldn't retrieve data. Error: {error}
                </MessageBar>);

        if (items === undefined || items.length <= 0)
            return (
                <div className="serviceStatusContent" style={{
                    textAlign: 'center',
                    justifyContent: 'center',
                    alignItems: 'center',
                    minWidth: '300px',
                    height: '267px',
                    overflowX: 'clip',
                    overflowY: 'auto'
                }}>
                    <FontIcon
                        iconName='SkypeCircleCheck'
                        className={iconClasses.Ok}
                    /><br /><br />
                    <Text variant='medium' style={{ position: 'relative', top: '-12px' }}><b>All services operational</b></Text>
                </div>
            );

        return (
            <div className="messageCenter7daysContent" style={{
                textAlign: 'center',
                justifyContent: 'center',
                alignItems: 'center',
                minWidth: '300px',
                height: '267px',
                overflowX: 'clip',
                overflowY: 'auto'
            }}>
                <DonutChart
                    culture={window.navigator.language}
                    data={{
                        chartTitle: 'Message Center items per service',
                        chartData: items,
                    }}
                    innerRadius={55}
                    legendsOverflowText={'More'}
                    width={287}
                    height={267}
                    hideLegend={false}
                    valueInsideDonut={items.reduce((partialSum, a) => partialSum + (a.data ? a.data : 0), 0)}
                /> 
        </div>
        );
    }

    componentDidMount() {
        this._getIncidents();
    }

    private _getIncidents() {
        var data: IChartDataPoint[] = [];
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
                    fetch('/api/items/statistics/currentWeek/messageCenter',
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
                            if (result && result.length > 0) {
                                var sortedResult = result.sort((a: any, b: any) => new Date(a.week) > new Date(b.week) ? -1 : 1);

                                var i = 0;

                                for (const item of sortedResult) {
                                    data.push({
                                        legend: item.service,
                                        data: item.items,
                                        color: _colors[0][i % _colors[0].length]
                                    });

                                    i++;
                                }
                            }

                            this.setState({
                                items: data,
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