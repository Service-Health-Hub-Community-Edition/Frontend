import * as React from 'react';
import { Fabric } from '@fluentui/react';
import { Text } from '@fluentui/react';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { FontIcon, mergeStyles, mergeStyleSets } from '@fluentui/react';
import { TooltipHost } from '@fluentui/react';
import { InfoCard } from '../InfoCard';
import { IServiceHealthEvent, IServiceStatistics } from '../IServiceStatistics';
import { ServiceComponent, ServiceComponentViewState } from '../ServiceNameComponent';
import { MessageBar, MessageBarType } from '@fluentui/react';
import { acquireAccessToken } from "../../auth/AccessTokenHelper";

interface IServiceStateCardCardState {
    statistics?: IServiceStatistics;
    isDataLoaded: boolean;
    error?: string;
}

export class ServiceStateCard extends React.Component<{ icon?: string, iconColor?: string, title?: string, href?: string, statistics?: IServiceStatistics }, IServiceStateCardCardState> {
    constructor(props: { icon?: string, iconColor?: string, title?: string, href?: string, statistics?: IServiceStatistics }) {
        super(props);

        this.state = {
            statistics: this.props.statistics,
            isDataLoaded: false,
            error: undefined
        };

    }

    public render() {
        const {
            isDataLoaded, statistics, error
        } = this.state;

        const iconClass = mergeStyles({
            fontSize: 60,
            height: 60,
            width: 60
        });

        const iconClasses = mergeStyleSets({
            Ok: [{ color: 'green' }, iconClass],
            Degradation: [{ color: 'orange' }, iconClass]
        });

        var serviceImpacted: boolean = statistics ?
            (statistics.eventStatistics.events.length > 0 ? true : false) : false;

        return (
            <Fabric>
                <InfoCard>
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

                    <div className="serviceStatistics" style={{
                        display: isDataLoaded && statistics !== undefined && error === undefined ? 'flex' : 'none',
                        justifyContent: 'center',
                        alignItems: 'center',
                        width: '100%',
                        height: '370px',
                        overflowX: 'clip',
                        overflowY: 'auto'
                    }}>
                        <div className="serviceStatisticsContent" style={{ textAlign: 'center' }}>
                            <FontIcon
                                iconName={serviceImpacted ? 'WarningSolid' : 'SkypeCircleCheck'}
                                className={serviceImpacted ? iconClasses.Degradation : iconClasses.Ok}
                            /><br /><br />
                            <Text variant='medium' style={{ position: 'relative', top: '-12px' }}><b>{serviceImpacted ? 'Service degradation' : 'All services operational'}</b></Text>
                            <br /><br />
                            <div className="serviceCount" style={{
                                textAlign: 'center', display: serviceImpacted ? '' : 'none' 
                            }}>
                                <Text variant='superLarge'>{statistics ? statistics.eventStatistics.impactedServices.length : '0'}</Text><br />
                                <Text variant='smallPlus' style={{ position: 'relative', top: '-12px'}}>impacted services</Text>
                                <br /><br /><br />
                                {statistics ? statistics.eventStatistics.impactedServices.map((t) => (
                                    <TooltipHost key={t} content={
                                        t + ': ' +
                                        statistics.eventStatistics.events
                                            .filter(e =>
                                                e.workload.toLowerCase().trim() === t.toLowerCase().trim() &&
                                                e.classification.toLowerCase().trim() === 'incident').length.toString() +
                                        (statistics.eventStatistics.events
                                            .filter(e =>
                                                e.workload.toLowerCase().trim() === t.toLowerCase().trim() &&
                                                e.classification.toLowerCase().trim() === 'incident').length === 1 ? ' incident, ' : ' incidents, ') +
                                        statistics.eventStatistics.events
                                            .filter(e => 
                                                e.workload.toLowerCase().trim() === t.toLowerCase().trim() &&
                                                e.classification.toLowerCase().trim() === 'advisory').length.toString() +
                                        (statistics.eventStatistics.events
                                            .filter(e =>
                                                e.workload.toLowerCase().trim() === t.toLowerCase().trim() &&
                                                e.classification.toLowerCase().trim() === 'advisory').length === 1 ? ' advisory' : ' advisories')
                                        }
                                        >
                                        <ServiceComponent name={t} viewState={ServiceComponentViewState.IconOnly}/>
                                    </TooltipHost>
                                )) : ""}
                            </div>
                        </div>
                    </div>
                </InfoCard>
            </Fabric>
        );
    }

    componentDidMount() {
        var statistics: IServiceStatistics | undefined = undefined;

        if (this.state.statistics === undefined) {
            acquireAccessToken()
                .then((response) => {
                    fetch('/api/Statistics', { headers: response.idToken === "" ? {} : { 'Authorization': `Bearer ${response.idToken}` } })
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
                            statistics = result;
                        }).then(() => {
                            this.setState({
                                statistics: statistics,
                                isDataLoaded: true
                            });
                        }).catch((err) => {
                            this.setState({
                                error: err.message
                            });
                        });
                }).catch((err) => {
                    this.setState({
                        error: err.message
                    });
                });
        }
        else {
            this.setState({
                isDataLoaded: true
            });
        }
    }
}