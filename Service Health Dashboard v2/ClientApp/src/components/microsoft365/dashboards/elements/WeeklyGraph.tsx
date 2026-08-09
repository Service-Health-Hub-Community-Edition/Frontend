import * as React from 'react';
import {
    Text, Spinner, SpinnerSize, 
    MessageBar, MessageBarType, ITheme
} from '@fluentui/react';
import { VerticalBarChart, IVerticalBarChartDataPoint } from '@fluentui/react-charting';
import { acquireAccessToken } from "../../../../auth/AccessTokenHelper";
import { checkInTeams } from '../../../auth/detectTeams';
import { AccessDenied } from "../../../AccessDenied";
import { GlobalState } from "../../../GlobalState";

export interface IWeeklyReportState {
    items: IVerticalBarChartDataPoint[];
    isDataLoaded: boolean;
    width: number;
    type: WeeklyReportType;
    container?: string;
    inTeams: boolean;
    error?: string;
    accessGranted?: boolean;
    reportContainer?: any;
}

export enum WeeklyReportType {
    MessageCenter = 'MessageCenter',
    ServiceHealth = 'ServiceHealth',
    Roadmap = 'Roadmap',
    AzureServiceHealth = 'AzureServiceHealth'
}

export class WeeklyGraph extends React.Component<{ type: WeeklyReportType, container?: string }, IWeeklyReportState> {
    static contextType: any = GlobalState;

    constructor(props: { type: WeeklyReportType, container?: string }) {
        super(props);

        this.state = {
            items: [],
            isDataLoaded: false,
            width: 800,
            type: props.type,
            container: props.container,
            inTeams: checkInTeams(),
            error: undefined,
            accessGranted: undefined,
            reportContainer: props.container ? document.getElementById(props.container) : undefined
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

        const { items, isDataLoaded, error, width } = this.state;

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
                    minWidth: '800px',
                    height: '267px',
                    overflowX: 'clip',
                    overflowY: 'auto'
                }}>
                    <Text variant='medium' style={{ position: 'relative', top: '-12px' }}><b>Weekly statistics are not available.</b></Text>
                </div>
            );

        return (
                <VerticalBarChart
                chartTitle="Weekly report"
                data={items}
                height={267}
                useSingleColor={true}
                yAxisTickCount={6}
                hideLegend={true}
                width={width}
                barWidth={width < 560 ? 20 : 32}
                rotateXAxisLables={width < 560}
                />
        );
    }

    componentDidMount() {
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

        this._getReport();
    }

    private getParentContainerWidth(): number {
        var width: number = 800;
        
        if (this.state.reportContainer && this.state.reportContainer !== null) {
            const rect = this.state.reportContainer?.getBoundingClientRect();
            width = rect.width;
        }

        return width;
    }

    private _getReport() {
        var data: IVerticalBarChartDataPoint[] = [];
        const requiredRoles: string[] = ['ServiceHealthReader', 'Communication.Write.All', 'Admin'];

        let globalState: any = this.context;
        var theme: ITheme = globalState.getTheme();

        acquireAccessToken()
            .then((response) => {
                var tokenClaims: any = response.account?.idTokenClaims;
                const userRoles: any = tokenClaims?.roles;
                var userHasRequiredRole: boolean = userRoles.some((r: string) => requiredRoles.includes(r));

                this.setState({
                    accessGranted: userHasRequiredRole
                });
                if (userHasRequiredRole)
                    fetch('/api/items/statistics/' + this.props.type.toString(),
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
                                        x: (new Date(item.week)).toLocaleDateString(window.navigator.language, { day: "numeric", month: "short" }),
                                        y: item.items,
                                        color: theme.semanticColors.accentButtonBackground
                                    });

                                    i++;
                                }
                            }

                            this.setState({
                                items: data,
                                type: this.props.type,
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