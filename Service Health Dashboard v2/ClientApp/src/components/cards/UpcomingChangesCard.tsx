import * as React from 'react';
import { Fabric } from '@fluentui/react';
import { Text, IFontStyles, ITheme } from '@fluentui/react';
import { Link } from '@fluentui/react';
import { Separator } from '@fluentui/react';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { InfoCard } from '../InfoCard';
import { StatisticsTile } from '../StatisticsTileComponent';
import { IServiceStatistics } from '../IServiceStatistics';
import { MessageBar, MessageBarType, ThemeProvider } from '@fluentui/react';
import { GlobalState } from '../GlobalState';
import { CountAnnotationBar, ICountAnnotationBarProps } from '@m365-admin/count-annotation';
import { acquireAccessToken } from "../../auth/AccessTokenHelper";

interface IUpcomingChangesCardState {
    icon?: string;
    iconColor: string;
    title?: string;
    href?: string;
    hrefTitle?: string;
    statistics?: IServiceStatistics;
    exoCount: number;
    spoCount: number;
    oneDriveCount: number;
    teamsCount: number;
    isDataLoaded: boolean;
    error?: string;
    theme?: any;
}

export class UpcomingChangesCard extends React.Component<{ icon?: string, iconColor?: string, title?: string, href?: string, hrefTitle?: string, statistics?: IServiceStatistics }, IUpcomingChangesCardState> {
    static contextType = GlobalState;

    constructor(props: { icon?: string, iconColor?: string, title?: string, href?: string, hrefTitle?: string, statistics?: IServiceStatistics }) {
        super(props);

        this.state = {
            icon: this.props.icon,
            iconColor: this.props.iconColor ? this.props.iconColor : "#2B2B2B",
            title: this.props.title,
            href: this.props.href,
            hrefTitle: this.props.hrefTitle,
            statistics: this.props.statistics,
            exoCount: 0,
            spoCount: 0,
            oneDriveCount: 0,
            teamsCount: 0,
            isDataLoaded: false,
            error: undefined
        };

    }

    public render() {
        const {
            icon, iconColor, title, href, hrefTitle, isDataLoaded, statistics, exoCount, spoCount, oneDriveCount, teamsCount, error, theme
        } = this.state;

        if (theme === undefined)
            return "";

        const messageCenter: ICountAnnotationBarProps = {
            countAnnotationProps: [
                {
                    annotationText: 'SharePoint Online',
                    count: spoCount,
                    annotationColor: '#0078d4'
                },
                {
                    annotationText: 'OneDrive',
                    count: oneDriveCount,
                    annotationColor: '#0078d4'
                },
                {
                    annotationText: 'Exchange Online',
                    count: exoCount,
                    annotationColor: '#0078d4'
                },
                {
                    annotationText: 'Microsoft Teams',
                    count: teamsCount,
                    annotationColor: '#6264a7'
                }
            ]
        };

        const roadmap: ICountAnnotationBarProps = {
            countAnnotationProps: [
                {
                    annotationText: 'In development',
                    count: statistics ? statistics.roadmapStatistics.inDevelopment : 0,
                    annotationColor: '#0078d4'
                },
                {
                    annotationText: 'Rolling out',
                    count: statistics ? statistics.roadmapStatistics.rollingOut : 0,
                    annotationColor: '#107c10'
                }
            ]
        };
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

                    <div className="upcomingChangesStats" style={{
                        display: isDataLoaded && statistics !== undefined && error === undefined ? 'block' : 'none',
                        width: '100%',
                        height: '330px',
                        overflowX: 'clip',
                        overflowY: 'clip'
                    }}>
                        <div className="mcHeader" style={{ textAlign: 'left', width: '100%', marginTop: '4px'}}>
                            <Text variant='small'><b>Message center</b></Text><br />
                        </div>

                        <br />
                        <div className="mcContent" style={{
                            display: 'flex',
                            marginBottom: '-14px',
                            alignItems: 'top'
                        }}>
                            
                                <CountAnnotationBar countAnnotationProps={messageCenter.countAnnotationProps} />

                            
                        </div>
                        <br />
                        <div className="mcStatisticsFooter" style={{ textAlign: 'right', width: '100%' }}>
                            <Link href='/messages' style={{ fontSize: 'x-small' }}>More...</Link>
                        </div>

                        <Separator />

                        <div className="roadmapHeader" style={{ textAlign: 'left', width: '100%' }}>
                            <Text variant='small'><b>Roadmap</b></Text><br />
                        </div>

                        <div className="roadmapContent" style={{
                            display: 'flex',
                            height: '72px',
                            alignItems: 'top'
                        }}>
                            
                                <CountAnnotationBar countAnnotationProps={roadmap.countAnnotationProps} />

                        </div>

                        <div className="rmStatisticsFooter" style={{ textAlign: 'right', width: '100%' }}>
                            <Link href='/roadmap' style={{ fontSize: 'x-small' }}>More...</Link>
                        </div>
                    </div>
                </InfoCard>
            </Fabric>
        );
    }

    componentDidMount() {
        var statistics: IServiceStatistics | undefined = undefined;
        var exo: number = 0;
        var spo: number = 0;
        var oneDrive: number = 0;
        var teams: number = 0;
        var statsLoaded: boolean = false;
        var mcStatsLoaded: boolean = false;

        let globalState: any = this.context;
        var theme: ITheme = globalState.getTheme();

        this.setState({
            theme: theme
        });

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
                            statsLoaded = true;
                        }).then(() => {
                            this.setState({
                                statistics: statistics,
                                isDataLoaded: statsLoaded && mcStatsLoaded
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
            statsLoaded = true;
            this.setState({
                isDataLoaded: statsLoaded && mcStatsLoaded
            });
        }

        acquireAccessToken()
            .then((response) => {
                fetch('/api/Messages?select=services&service=Exchange Online, SharePoint Online, Microsoft Teams, OneDrive for Business',
                    { headers: response.idToken === "" ? {} : { 'Authorization': `Bearer ${response.idToken}` } })
                    .then(response => {
                        if (response.ok) {
                            return response.json();
                        } else {
                            var errorMsg: string = this.state.error === undefined ? "" : this.state.error + "<br/>";
                            this.setState({
                                error: errorMsg + "Fetching data from /api/Messages endpoint failed with an error " + response.status + " " + response.statusText
                            });
                            throw Error(response.status + " " + response.statusText);
                        }
                    })
                .then(result => {
                    exo = this._getFilteredItemsCount(result, "Exchange Online");
                    spo = this._getFilteredItemsCount(result, "SharePoint Online");
                    teams = this._getFilteredItemsCount(result, "Microsoft Teams");
                    oneDrive = this._getFilteredItemsCount(result, "OneDrive for Business");

                    mcStatsLoaded = true;
                }).then(() => {
                    this.setState({
                        exoCount: exo,
                        spoCount: spo,
                        teamsCount: teams,
                        oneDriveCount: oneDrive,
                        isDataLoaded: statsLoaded && mcStatsLoaded
                    });
                });
            }).catch((err) => {
                var errorMsg: string = this.state.error === undefined ? "" : this.state.error + "<br/>";
                this.setState({
                    error: errorMsg + "Failed retrieving data from /api/Messages with an error " + err.message
                });
            });
    }

    private _getFilteredItemsCount(result: any, service: string): number {
        var items = result.filter(
            (i: any) => {
                var services = i.services.find(
                    (w: any) => w.trim().toLowerCase() === service.toLowerCase());
                return services !== undefined;
            });

            return (items ? items.length : 0);
    }
}