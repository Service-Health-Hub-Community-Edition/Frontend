import * as React from 'react';
import { Fabric } from '@fluentui/react';
import { Text } from '@fluentui/react';
import { Link } from '@fluentui/react';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { IImageProps, Image, ImageFit } from '@fluentui/react';
import { FontIcon, mergeStyles, mergeStyleSets } from '@fluentui/react';
import {
    DetailsList, SelectionMode, DetailsListLayoutMode,
    IColumn, ILinkStyles, ILinkStyleProps,
    IDetailsColumnProps, IDetailsColumnStyles
} from '@fluentui/react';
import { DefaultButton, PrimaryButton, IconButton, Panel, PanelType, Modal } from '@fluentui/react';
import { InfoCard } from '../InfoCard';
import { IServiceHealthEvent, IServiceStatistics } from '../IServiceStatistics';
import { IncidentDetails } from '../IncidentDetails';
import { MessageBar, MessageBarType } from '@fluentui/react';
import { acquireAccessToken } from "../../auth/AccessTokenHelper";

interface IActiveEventsCardState {
    icon?: string;
    iconColor: string;
    title?: string;
    href?: string;
    hrefTitle?: string;
    statistics?: IServiceStatistics;
    isDataLoaded: boolean;
    isMessageDetailsDialogOpen: boolean;
    selectedMessage: string;
    error?: string;
}

export class ActiveEventsCard extends React.Component<{ icon?: string, iconColor?: string, title?: string, href?: string, hrefTitle?: string, statistics?: IServiceStatistics  }, IActiveEventsCardState> {
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
            isMessageDetailsDialogOpen: false,
            selectedMessage: "",
            error: undefined
        };

    }

    handleItemChange = (itemId: string, published: boolean) => {
        if (this.state.statistics) {
            var item = this.state.statistics.eventStatistics.events.find(m => m.id.toLowerCase() === itemId.toLowerCase());
            if (item !== null && item !== undefined) {
                item.published = published;
            }
        }
    }

    public render() {
        const {
            icon, iconColor, title, href, hrefTitle, isDataLoaded, statistics, isMessageDetailsDialogOpen, selectedMessage, error
        } = this.state;

        const pipeFabricStyles = (p: ILinkStyleProps): ILinkStyles => ({
            root: {
                textDecoration: 'none',
                color: p.theme.semanticColors.bodyText,
                fontWeight: '600',
                fontSize: p.theme.fonts.small.fontSize,
            },
        });

        const iconClass = mergeStyles({
            fontSize: 16,
            height: 16,
            width: 16,
            margin: "0 16px"
        });

        const classNames = mergeStyleSets({
            deepSkyBlue: [{ color: 'deepskyblue' }, iconClass],
            greenYellow: [{ color: 'greenyellow' }, iconClass],
            salmon: [{ color: 'salmon' }, iconClass],
            black: [{ color: 'black' }, iconClass],
        });

        const columns: IColumn[] = [
            {
                key: 'clId',
                name: 'Id',
                fieldName: 'id',
                minWidth: 60,
                maxWidth: 60,
                isRowHeader: true,
                isResizable: false,
                data: 'string',
                isPadded: true,
                onRender: (item: IServiceHealthEvent) => {
                    if (item.published)
                    {
                        return <div>
                            <Text variant='small'>{item.id}</Text>
                            <FontIcon aria-label="Published" iconName="News" className={classNames.black} />
                        </div>;
                    }
                    else
                        return <Text variant='small'>{item.id}</Text>;
                },
            },
            {
                key: 'clTitle',
                name: 'Title',
                fieldName: 'title',
                minWidth: 100,
                maxWidth: 100,
                isRowHeader: true,
                isResizable: true,
                isMultiline: true,
                onRender: (item: IServiceHealthEvent) => {
                    return <Link onClick={(event) => {
                        event.preventDefault();
                        this._onOpenMessageDetails(item.id);
                    }} styles={pipeFabricStyles}>{item.title}</Link >;
                },
                data: 'string',
                isPadded: true,
            }
        ];

        var serviceHealthEvents: IServiceHealthEvent[] = statistics ? statistics.eventStatistics.events : [];

        return (
            <Fabric>
            <InfoCard
                    icon={icon}
                    iconColor={iconColor}
                    title={title}
                    href={href}
                    hrefTitle={hrefTitle}
            >
                <div className="loadingProgress" style={{ display: isDataLoaded || error !== undefined ? 'none' : 'block' }}>
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

                <div className="activeEventsList" style={{
                        display: isDataLoaded && statistics !== undefined && error === undefined ? 'block' : 'none',
                        width: '100%',
                        height: '330px',
                        overflowX: 'clip',
                        overflowY: 'auto'
                }}>
                        <DetailsList
                            items={serviceHealthEvents}
                            compact={false}
                            columns={columns}
                            selectionMode={SelectionMode.none}
                            layoutMode={DetailsListLayoutMode.justified}
                            isHeaderVisible={false}
                        />

                        <Panel
                            isOpen={isMessageDetailsDialogOpen}
                            onDismiss={this._onDismisMessageDetails}
                            headerText={statistics ? statistics.eventStatistics.events.find(m => m.id.toLowerCase() === selectedMessage.toLowerCase())?.title! : ""}
                            type={PanelType.medium}
                            isLightDismiss
                            onRenderNavigationContent={(props, defaultRender) => (
                                <div>
                                    <IconButton
                                        iconProps={{ iconName: 'Up' }}
                                        title='Previous'
                                        ariaLabel='Previous item'
                                        disabled={this._isFirst(selectedMessage)}
                                        onClick={() => this._selectPrevious(selectedMessage)} />

                                    <IconButton
                                        iconProps={{ iconName: 'Down' }}
                                        title='Next'
                                        ariaLabel='Next item'
                                        disabled={this._isLast(selectedMessage)}
                                        onClick={() => this._selectNext(selectedMessage)} />

                                    {defaultRender!(props)}
                                </div>
                            )
                            }
                        >
                            <IncidentDetails id={selectedMessage} onPublishingChange={this.handleItemChange} />
                        </Panel>
                </div>

                <div className="noActiveIncidents" style={{
                    display: isDataLoaded && statistics === undefined ? 'flex' : 'none',
                    justifyContent: 'center',
                    alignItems: 'center',
                    width: '100%',
                    height: '300px'
                }}>
                    <div className="noActiveIncidentsContents" style={{ textAlign: 'center' }}>
                        <img src='/images/well-done.svg' width='140px' /><br /><br />
                        <Text variant='medium'>There are no active service events!</Text><br />
                        <Link href='/'>Open service event history</Link>
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

    private _onOpenMessageDetails(id: string): void {
        this.setState({
            selectedMessage: id,
            isMessageDetailsDialogOpen: true
        });
    }

    private _onDismisMessageDetails = (): void => {
        this.setState({
            selectedMessage: "",
            isMessageDetailsDialogOpen: false
        });
    }

    private _selectPrevious(id: string): void {
        if (this.state.statistics !== undefined) {
            const itemIndex: number = this.state.statistics.eventStatistics.events.findIndex((i) => i.id === id);

            if (itemIndex > 0) {
                this.setState({
                    selectedMessage: this.state.statistics.eventStatistics.events[itemIndex - 1].id
                });
            }
        }
    }

    private _selectNext(id: string): void {
        if (this.state.statistics !== undefined) {
            const itemIndex: number = this.state.statistics.eventStatistics.events.findIndex((i) => i.id === id);

            if (itemIndex < this.state.statistics.eventStatistics.events.length) {
                this.setState({
                    selectedMessage: this.state.statistics.eventStatistics.events[itemIndex + 1].id
                });
            }
        }
    }

    private _isFirst(id: string): boolean {
        if (this.state.statistics !== undefined) {
            const firstElement: IServiceHealthEvent = this.state.statistics.eventStatistics.events[0];
            return firstElement.id === id;
        } else
            return true;
    }

    private _isLast(id: string): boolean {
        if (this.state.statistics !== undefined) {
            const lastElement: IServiceHealthEvent = this.state.statistics.eventStatistics.events[this.state.statistics.eventStatistics.events.length - 1];
            return lastElement.id === id;
        } else
            return true;
    }
}