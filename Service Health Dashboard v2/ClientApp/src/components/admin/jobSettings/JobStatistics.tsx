import * as React from 'react';
import {
    Text, Spinner, SpinnerSize, DefaultPalette, mergeStyles
} from '@fluentui/react';
import { HorizontalBarChart, IChartProps } from '@fluentui/react-charting';
import * as d3 from 'd3-format';
import { acquireAccessToken } from "../../../auth/AccessTokenHelper";
import { AuthenticationResult } from '@azure/msal-browser';
import { IJobHistory } from './JobHistory';

interface IJobStatisticsState {
    componentId: string;
    items: IJobHistory[];
    accessGranted: boolean;
    initialized: boolean;
    error?: string;
}

const calloutItemStyle = mergeStyles({
    borderBottom: '1px solid #D9D9D9',
    padding: '3px',
});

export class JobStatistics extends React.Component<{ componentId: string }, IJobStatisticsState> {

    constructor(props: { componentId: string }) {
        super(props);

        this.state = {
            componentId: props.componentId,
            items: [],
            accessGranted: false,
            initialized: false
        }
    }

    render() {
        const { items, accessGranted, initialized, error } = this.state;

        if (!initialized)
            return (<>
                <br /><br />
                <Spinner size={SpinnerSize.medium} />
            </>);

        if (error !== undefined)
            return (<>
                <br /><br />
                <Text variant='medium'><b>{error}</b></Text>
            </>);

        if (!accessGranted)
            return (<>
                <br /><br />
                <Text variant='medium'><b>You have no access to this ressource.</b></Text>
            </>);

        if (!items || items.length <= 0)
            return (<>
                <br /><br />
                <Text variant='medium'><b>No events in past 24 hours.</b></Text>
            </>);

        const data: IChartProps[] = [
            {
                chartTitle: 'Completed',
                chartData: [
                    {
                        legend: 'Completed jobs',
                        horizontalBarChartdata: { x: items.filter((i: IJobHistory) => i.state === 'Completed')?.length, y: items.length },
                        color: DefaultPalette.green,
                    },
                ],
            },
            {
                chartTitle: 'Failed',
                chartData: [
                    {
                        legend: 'Failed jobs',
                        horizontalBarChartdata: { x: items.filter((i: IJobHistory) => i.state === 'Failed')?.length, y: items.length },
                        color: DefaultPalette.redDark,
                    },
                ],
            }
        ];

        return (<>
            <HorizontalBarChart culture={window.navigator.language} data={data} hideRatio={[true]} />
        </>);
    }

    componentDidMount() {
        this.setState({
            componentId: this.props.componentId
        });

        this._loadData();
    }

    _loadData(): void {
        const requiredRoles: string[] = ['Admin'];
        var authResponse: AuthenticationResult;
        var userHasRequiredRole: boolean = false;

        this.setState({
            initialized: false
        });

        acquireAccessToken()
            .then((response) => {
                var tokenClaims: any = response.account?.idTokenClaims;
                const userRoles: any = tokenClaims?.roles;
                userHasRequiredRole = userRoles.some((r: string) => requiredRoles.includes(r));

                this.setState({
                    accessGranted: userHasRequiredRole
                });

                authResponse = response;
            })
            .then(() => {
                if (userHasRequiredRole)
                    fetch('/api/jobs/' + this.state.componentId + '/history', { headers: authResponse.idToken === "" ? {} : { 'Authorization': `Bearer ${authResponse.idToken}` } })
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
                    .then(response => {
                        this.setState({
                            items: response.filter((i: IJobHistory) => i.start && (new Date(i.start) >= new Date(new Date().setHours(-24)))),
                            initialized: true
                        });
                    })
            }).catch((err) => {
                this.setState({
                    error: err.message
                });
            });
    }
}