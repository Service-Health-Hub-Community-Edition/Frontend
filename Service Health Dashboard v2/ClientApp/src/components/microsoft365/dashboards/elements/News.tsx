import * as React from 'react';
import { ColumnActionsMode, DefaultPalette } from '@fluentui/react';
import { Text, Link, ILinkStyleProps, ILinkStyles, TooltipHost, ScrollablePane, Sticky, IDetailsHeaderProps, IRenderFunction } from '@fluentui/react';
import {
    IColumn, DetailsList, SelectionMode, DetailsListLayoutMode, ConstrainMode, FontIcon,
    IDetailsListStyles, mergeStyles, mergeStyleSets,
    HoverCard, IExpandingCardProps, DirectionalHint, IconButton, IContextualMenuProps
} from '@fluentui/react';
import { Panel, PanelType } from '@fluentui/react';
import { Tag, TagList, TagType, ITagProps } from '@m365-admin/tag';
import { MessageDetails } from '../../../MessageDetails';
import { ICustomAction, CustomAction } from '../../../CustomAction';
import { ServiceComponent } from '../../../ServiceNameComponent';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { MessageBar, MessageBarType } from '@fluentui/react';
import { acquireAccessToken, accessControl } from "../../../../auth/AccessTokenHelper";
import { checkInTeams } from '../../../auth/detectTeams';
import { AccessDenied } from "../../../AccessDenied";
import { ISummary, ISummaryElement } from "./Incidents";
import { setViewState, setArchiveState, setFavoriteState } from "../../../../api/viewpoint";

export interface INewsEvent {
    id: string;
    title: string;
    service: string[];
    tags: string[];
    category: string;
    majorUpdate: boolean;
    published: Date;
    public?: boolean;
    lastModified: Date;
    summary: ISummary;
    joinedTags: string[];
    rolloutStatus?: string;
    serviceHealthHubState: string;
    serviceHealthHubViewpoint: any;
}

interface INewsState {
    items: INewsEvent[];
    isDetailsPanelOpen: boolean;
    selectedItem: string;
    isDataLoaded: boolean;
    inTeams: boolean;
    customActions: ICustomAction[] | undefined;
    error?: string;
    accessGranted?: boolean;
}


const iconClass = mergeStyles({

});

const classNames = mergeStyleSets({
    incident: [{ color: 'red' }, iconClass],
    advisory: [{ color: DefaultPalette.blue }, iconClass],
    compactCard: {
        display: 'flex',
        height: '100%',
    },
    expandedCard: {
        padding: '16px 24px',
    },
    item: {
        selectors: {
            '&:hover': {
                textDecoration: 'underline',
                cursor: 'pointer',
            },
        },
    }
});

export enum NewsItems {
    All = 0,
    Inbox = 1,
    Archived = 2,
    Favorites = 3
}

const componentName: string = 'ServiceUpdateMessage';

export class News extends React.Component<{ collectionMode?: NewsItems }, INewsState> {
    _allNews: INewsEvent[] = [];
    customAction: any = React.createRef();

    constructor(props: { collectionMode?: NewsItems }) {
        super(props);

        this.state = {
            items: [],
            isDetailsPanelOpen: false,
            selectedItem: "",
            isDataLoaded: false,
            inTeams: checkInTeams(),
            error: undefined,
            accessGranted: undefined,
            customActions: undefined
        };
    }

    handleItemChange = (itemId: string, published: boolean) => {
        var item = this._allNews.find(m => m.id.toLowerCase() === itemId.toLowerCase());
        if (item !== null && item !== undefined) {
            item.public = published;
            this.setState({
                items: this._getViewTypeFilter(this._allNews)
            });
        }
    }

    handleView = (itemId: string, viewed: boolean) => {
        var item = this._allNews.find(m => m.id.toLowerCase() === itemId.toLowerCase());
        if (item !== null && item !== undefined && item.serviceHealthHubViewpoint !== undefined) {
            item.serviceHealthHubViewpoint.viewed = viewed;
            this.setState({
                items: this._getViewTypeFilter(this._allNews)
            });
        }
    }

    handleFavorite = (itemId: string, favorite: boolean) => {
        var item = this._allNews.find(m => m.id.toLowerCase() === itemId.toLowerCase());
        if (item !== null && item !== undefined && item.serviceHealthHubViewpoint !== undefined) {
            item.serviceHealthHubViewpoint.favorite = favorite;
            this.setState({
                items: this._getViewTypeFilter(this._allNews)
            });
        }
    }

    handleArchive = (itemId: string, archived: boolean) => {
        var item = this._allNews.find(m => m.id.toLowerCase() === itemId.toLowerCase());
        if (item !== null && item !== undefined && item.serviceHealthHubViewpoint !== undefined) {
            item.serviceHealthHubViewpoint.archived = archived;
            this.setState({
                items: this._getViewTypeFilter(this._allNews)
            });
        }
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

        const { items, isDetailsPanelOpen, selectedItem, isDataLoaded, customActions, error } = this.state;

        const pipeFabricStyles = (p: ILinkStyleProps): ILinkStyles => ({
            root: {
                textDecoration: 'none',
                color: p.theme.semanticColors.bodyText
            },
        });

        const onRenderCompactCard = (item: INewsEvent): JSX.Element => {
            return (
                <div className={classNames.compactCard}>
                    <div className="container" style={{ margin: "12px" }} >
                        <div className="row" style={{ paddingBottom: "6px" }} >
                            <div className="col-auto" style={{marginRight: "0px", paddingRight: "6px"}} >
                                <FontIcon
                                    iconName={item.majorUpdate ? 'WarningSolid' : 'InfoSolid'}
                                    className={item.majorUpdate ? classNames.incident : classNames.advisory} />
                            </div>
                            <div className="col" style={{ margin: "0px", padding: "0px", alignItems: 'center' }}>
                                <Text variant='smallPlus'><b>{item.majorUpdate ? "Major update, " + item.category : item.category}</b></Text>
                            </div>
                        </div>
                        <div className="row" style={{ paddingBottom: "6px" }} >
                            <div className="col">
                                <Text variant='small'><b>{item.title}</b></Text>
                            </div>
                        </div>
                        <div className="row">
                            <div className="col" >
                                <Text variant='xSmall'>
                                    {item.id} · Published {item.published.toLocaleDateString([], { year: 'numeric', month: 'short', day: 'numeric' })} · Updated {item.lastModified.toLocaleDateString([], { year: 'numeric', month: 'short', day: 'numeric' })}
                                </Text>
                            </div>
                        </div>
                    </div>
                </div>
            );
        };

        const onRenderExpandedCard = (item: INewsEvent): JSX.Element => {
            var iconName: string | undefined = undefined;
            switch (item.rolloutStatus?.toLowerCase()) {
                case 'rolling out':
                    iconName = 'Airplane';
                    break;
                case 'scheduled':
                    iconName = 'Clock';
                    break;
                case 'launched':
                    iconName = 'Completed';
                    break;
                default:
                    iconName = 'Info';
                    break;
            }

            return (
                <div className={classNames.expandedCard}>
                    {item.rolloutStatus || item.public || !(item.serviceHealthHubState == null || item.serviceHealthHubState == undefined || item.serviceHealthHubState.trim() == "") ? (<>
                        <div style={{ marginBottom: '6px' }}>
                            { renderStatusTag(item.id, item.serviceHealthHubState) }
                            &nbsp;
                            {item.public ? (<Tag
                                onRenderContent={(tagProps: ITagProps | undefined, defaultRender: IRenderFunction<ITagProps> | undefined): JSX.Element => {
                                    return <><FontIcon iconName='Streaming' /> {defaultRender && tagProps ? defaultRender(tagProps) : ""}</>;
                                }} >
                                Published
                            </Tag>) : (<></>)}
                            &nbsp;
                            {item.rolloutStatus ? (<Tag
                                onRenderContent={(tagProps: ITagProps | undefined, defaultRender: IRenderFunction<ITagProps> | undefined): JSX.Element => {
                                    return <><FontIcon iconName={iconName} /> {defaultRender && tagProps ? defaultRender(tagProps) : ""}</>;
                                }} >
                                {item.rolloutStatus}
                            </Tag>) : (<></>)}
                        </div></>
                    ) : ""}
                    {item.summary?.contents && item.summary?.contents.length > 0 ? (
                        <>
                            <Text variant='medium'><b>Summary</b></Text><br />
                            { item.summary.timestamp ? (
                                <Text variant='xSmall'>Updated: {new Date(item.summary.timestamp).toLocaleString()}</Text>) : ""
                            }
                            {item.summary?.contents.length > 1 ? (
                                <ul style={{ paddingTop: "6px" }} >
                                    {item.summary.contents.map((summaryLine: any) => (
                                        <li style={{ paddingTop: "6px" }}><Text variant='small'>{summaryLine.text}</Text></li>
                                    ))}
                                </ul>) : (<div><br />
                                    <Text variant='small'>{item.summary.contents[0].text}</Text></div>
                            )}
                        </>
                    ) : (
                            <div style={{ textAlign: 'center'}}>
                                <Text variant='medium'>No summary available</Text>
                            </div>
                    )}
                </div>
            );
        };

        const renderStatusTag = (id: string, status: string | null | undefined) => {
            if (status == null || status == undefined || status == "")
                return <></>

            var tagType: TagType | undefined;
            switch (status.trim().toUpperCase()) {
                case "NEW":
                    tagType = TagType.New;
                    break;
                case "UPDATED":
                    tagType = TagType.HighImpactInformational;
                    break;
                default:
                    tagType = TagType.LowImpactInformational;
                    break;
            }

            return <>
                <Tag
                    key={id + '-state'}
                    tagType={tagType}
                >
                    {status}
                </Tag>
            </>
        }

        const stateEmpty: boolean = items.find((i: INewsEvent) => !(i.serviceHealthHubState == null || i.serviceHealthHubState == undefined || i.serviceHealthHubState == "")) == undefined;

        const columns: IColumn[] = [
            {
                key: 'clType',
                name: '',
                minWidth: 4,
                maxWidth: 4,
                isResizable: false,
                isCollapsible: false,
                isMultiline: false,
                currentWidth: 4,
                onRender: (item: INewsEvent) => {
                    const expandingCardProps: IExpandingCardProps = {
                        onRenderCompactCard: onRenderCompactCard,
                        onRenderExpandedCard: onRenderExpandedCard,
                        renderData: item,
                        directionalHint: DirectionalHint.leftTopEdge,
                    };

                    return <div className="container" style={{ cursor: 'default' }} >
                        <HoverCard expandingCardProps={expandingCardProps}>
                            <div className="row" >
                                <div className="col" style={{ justifyContent: 'center', paddingLeft: '0px', paddingRight: '0px', marginRight: '3px' }} >
                                    {item.public ? (<FontIcon iconName='Streaming' />) : ""}
                                </div>
                                <div className="col" style={{ justifyContent: 'center', paddingLeft: '0px', paddingRight: '3px' }} >
                                    <FontIcon
                                        iconName={item.majorUpdate ? 'WarningSolid' : 'InfoSolid'}
                                        className={item.majorUpdate ? classNames.incident : classNames.advisory} />
                                </div>
                            </div>
                        </HoverCard></div>;
                },
                isPadded: true,
            },
            {
                key: 'clTitle',
                name: 'Title',
                fieldName: 'title',
                minWidth: 180,
                maxWidth: 350,
                isResizable: true,
                isCollapsible: false,
                isMultiline: false,
                data: 'string',
                onRender: (item: INewsEvent) => {
                    return (<Link onClick={(event) => {
                                    event.preventDefault();
                                    this._onOpenDetailsPanel(item.id);
                    }} styles={pipeFabricStyles}>{
                            item.serviceHealthHubViewpoint && item.serviceHealthHubViewpoint.viewed ?
                                item.title :
                                (<b>{item.title}</b>)
                        }</Link >);
                },
                isPadded: false,
            },
            {
                key: 'clSHHItemMenu',
                name: '',
                minWidth: 8,
                maxWidth: 8,
                isResizable: false,
                isCollapsible: true,
                isPadded: false,
                styles: {
                    
                    root: { padding: '6px 0px 6px 0px'}
                },
                onRender: (item: INewsEvent) => {
                    var menuProps: IContextualMenuProps = {
                        directionalHintFixed: true,
                        items: [
                            {
                                key: item.id + '-markAsRead',
                                text: item.serviceHealthHubViewpoint && item.serviceHealthHubViewpoint.viewed ? 'Mark as unread' : 'Mark as read',
                                iconProps: { iconName: item.serviceHealthHubViewpoint && item.serviceHealthHubViewpoint.viewed ? 'Read' : 'Mail' },
                                onClick: () => this._switchViewState(item.id)
                            },
                            {
                                key: item.id + '-archive',
                                text: item.serviceHealthHubViewpoint && item.serviceHealthHubViewpoint.archived ? 'Restore' : 'Archive',
                                iconProps: { iconName: item.serviceHealthHubViewpoint && item.serviceHealthHubViewpoint.viewed ? 'Undo' : 'Archive' },
                                onClick: () => this._switchArchiveState(item.id)
                            },
                            {
                                key: item.id + '-favorites',
                                text: item.serviceHealthHubViewpoint && item.serviceHealthHubViewpoint.favorited ? 'Remove from favorites' : 'Add to favorites',
                                iconProps: { iconName: item.serviceHealthHubViewpoint && item.serviceHealthHubViewpoint.favorited ? 'FavoriteStarFill' : 'FavoriteStar' },
                                onClick: () => this._switchFavoriteState(item.id)
                            }
                        ]
                    };

                    if (customActions && customActions.length > 0)
                        menuProps.items.push({
                            key: item.id + '-integrate',
                            text: 'Integrate',
                            iconProps: { iconName: 'AppIconDefault' },
                            subMenuProps: customActions ? {
                                items: customActions.map((ca: ICustomAction) => ({
                                    key: item.id + '-' + ca.actionId,
                                    text: ca.name,
                                    iconProps: ca.icon ? { iconName: ca.icon } : undefined,
                                    onClick: () => this.customAction.current._onRunCustomAction(ca.name, ca.actionId, item.id),
                                }))
                            } : undefined
                        });

                    return (<>
                        <IconButton
                            style={{ height: '16px' }}
                            iconProps={{ iconName: 'MoreVertical' }}
                            menuProps={menuProps}
                            menuIconProps={{ hidden: true }} />
                    </>)
                }
            },
            {
                key: 'clSHHState',
                name: '',
                minWidth: stateEmpty ? 1 : 64,
                maxWidth: stateEmpty ? 1 : 64,
                isResizable: false,
                isCollapsible: true,
                isPadded: false,
                onRender: (item: INewsEvent) => {
                    return renderStatusTag(item.id, item.serviceHealthHubState)
                },
            },
            {
                key: 'clServices',
                name: 'Affected services',
                minWidth: 100,
                maxWidth: 100,
                isResizable: false,
                isCollapsible: true,
                isPadded: false,
                onRender: (item: INewsEvent) => {
                    return <>
                        <TooltipHost content={item.service.join(", ")}><span>{
                            item.service.map((s) => (
                                <ServiceComponent name={s} />
                            ))
                        }</span></TooltipHost>
                        </>;
                },
            },
            {
                key: 'clRolloutStatus',
                name: '',
                minWidth: 4,
                maxWidth: 4,
                isResizable: false,
                isCollapsible: false,
                isMultiline: false,
                currentWidth: 4,
                onRender: (item: INewsEvent) => {
                    var iconName: string | undefined = undefined;
                    switch (item.rolloutStatus?.toLowerCase()) {
                        case 'rolling out':
                            iconName = 'Airplane';
                            break;
                        case 'scheduled':
                            iconName = 'Clock';
                            break;
                        case 'launched':
                            iconName = 'Completed';
                            break;
                        default:
                            iconName = 'Info';
                            break;
                    }

                    return item.rolloutStatus ? (
                        <TooltipHost content={'Status for your org: ' + item.rolloutStatus}>
                            <div className="container" style={{ cursor: 'default' }} >
                                <div className="row" >
                                    <div className="col" style={{ justifyContent: 'center' }} >
                                        <FontIcon
                                            iconName={iconName}
                                            className={classNames.advisory} />
                                    </div>
                                </div>
                            </div>
                        </TooltipHost>) : (<></>);
                },
                isPadded: true,
            },
            {
                key: 'clUpdated',
                name: 'Updated',
                minWidth: 60,
                maxWidth: 60,
                isResizable: false,
                isCollapsible: true,
                isPadded: false,
                onRender: (item: INewsEvent) => {
                    return <>{item.lastModified ? item.lastModified.toLocaleDateString() : ""}</>;
                },
            },
            {
                key: 'clTags',
                name: 'Tags',
                fieldName: 'tags',
                minWidth: 120,
                maxWidth: 120,
                isResizable: true,
                isCollapsible: true,
                data: 'string',
                isPadded: false,
                onRender: (item: INewsEvent) => {
                   return item.joinedTags ? (<TagList tags={item.joinedTags.map((tag: string) => ({ key: tag, children: tag, }))} />) : ""
                },
            },
            {
                key: 'clId',
                name: 'Id',
                minWidth: 80,
                maxWidth: 80,
                isResizable: true,
                isCollapsible: true,
                data: 'string',
                isPadded: false,
                onRender: (item: INewsEvent) => {          
                    return <>{item.id}</>;
                },
            }
        ];

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

        return (<>
                <CustomAction componentName={componentName} ref={this.customAction} onLoad={(actions: ICustomAction[]) => this._onLoadCustomActions(actions)} />

                <ScrollablePane style={{
                    top: '64px',
                    height: '260px'
                }} >
                <DetailsList
                    items={items}
                    compact={true}
                    columns={columns}
                    styles={{
                        contentWrapper: {
                            overflowX: 'hidden'
                        },
                        root: {
                            overflowX: 'hidden'
                        }
                    }}
                    selectionMode={SelectionMode.none}
                    layoutMode={DetailsListLayoutMode.justified}
                    constrainMode={ConstrainMode.horizontalConstrained}
                    isHeaderVisible={true}
                    onRenderDetailsHeader={
                        // tslint:disable-next-line:jsx-no-lambda
                        (detailsHeaderProps: IDetailsHeaderProps | undefined, defaultRender: IRenderFunction<IDetailsHeaderProps> | undefined) => (
                            <Sticky >
                                {defaultRender ? defaultRender(detailsHeaderProps) : ""}
                            </Sticky>
                        )}
                    />
            </ScrollablePane>

            <Panel
                headerText={items.find(m => m.id.toLowerCase() === selectedItem.toLowerCase())?.title!}
                isOpen={isDetailsPanelOpen}
                onDismiss={this._onDismisDetailsPanel}
                type={PanelType.medium}
                // You MUST provide this prop! Otherwise screen readers will just say "button" with no label.
                closeButtonAriaLabel="Close"
                hasCloseButton={false}
                    onRenderNavigationContent={(props, defaultRender) => (
                        <div>
                            <IconButton
                                iconProps={{ iconName: 'Up' }}
                                title='Previous'
                                ariaLabel='Previous item'
                                disabled={this._isFirst(selectedItem)}
                                onClick={() => this._selectPrevious(selectedItem)} />

                            <IconButton
                                iconProps={{ iconName: 'Down' }}
                                title='Next'
                                ariaLabel='Next item'
                                disabled={this._isLast(selectedItem)}
                                onClick={() => this._selectNext(selectedItem)} />

                            <IconButton
                                iconProps={{ iconName: 'Cancel' }}
                                title='Close'
                                ariaLabel='Close panel'
                                onClick={() => this._onDismisDetailsPanel()} />

                            {defaultRender!(props)}
                        </div>
                    )
                    }
            >
                <MessageDetails
                    id={selectedItem}
                    onPublishingChange={this.handleItemChange}
                    onView={this.handleView}
                    onArchive={this.handleArchive}
                    onFavorite={this.handleFavorite}
                />
            </Panel>
        </>
        );
    }

    componentDidMount() {
        this._getNews();
    }

    private _getNews() {
        var news: INewsEvent[] = [];
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
                    // fetch('/api/messages?select=id,title,services,details,startDateTime,lastModifiedDateTime,isMajorChange,categoryDisplayName,tags,public,serviceHealthHubState&expand=summary,serviceHealthHubViewpoint&filter=lastModifiedDateTime ge ' + (new Date(new Date().setDate(new Date().getDate() - 7)).toISOString()),
                    fetch('/api/messages?filter=lastModifiedDateTime ge ' + (new Date(new Date().setDate(new Date().getDate() - 7)).toISOString()),
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
                                var sortedResult = result.sort((a: any, b: any) => new Date(a.lastModifiedDateTime) > new Date(b.lastModifiedDateTime) ? -1 : 1);

                                for (const item of sortedResult) {
                                    var start: Date = new Date(item.startDateTime);
                                    var lastModified: Date = new Date(item.lastModifiedDateTime);

                                    var featureStatusJson = item.details.find((d: any) => d.name === 'FeatureStatusJson');
                                    var rolloutStatus: string | undefined = undefined;

                                    if (featureStatusJson !== undefined) {
                                        try {
                                            var featureStatus = JSON.parse(featureStatusJson.value);
                                            var fsKeys = Object.keys(featureStatus)
                                            var globalStateRoadmap = featureStatus[fsKeys[0]].find((gsr: any) => gsr.Platform === 'All');

                                            rolloutStatus = globalStateRoadmap && globalStateRoadmap.Status !== 'FeatureRolloutStatusNotSupported' ? globalStateRoadmap.Status : undefined
                                            if (rolloutStatus?.toLowerCase() === 'inrollout')
                                                rolloutStatus = 'Rolling out'
                                        } catch {
                                            rolloutStatus = undefined
                                        }
                                    }

                                    var orgTags: string[] = item.organizationTags.map((ot: any) => ot.tag);
                                    var tags: string[] = item.tags ? item.tags.sort((a: string, b: string) => (a > b ? 1 : -1)) : [];
                                    var joinedTags: string[] = [];

                                    for (const t of tags)
                                        joinedTags.push(t);

                                    for (const t of orgTags)
                                        joinedTags.push('[ORG] ' + t);

                                    news.push({
                                        id: item.id,
                                        title: item.title,
                                        service: item.services,
                                        published: start,
                                        lastModified: lastModified,
                                        category: item.categoryDisplayName,
                                        majorUpdate: item.isMajorChange,
                                        tags: item.tags ? item.tags.sort((a: any, b:any) => a > b ? 1 : -1) : [],
                                        public: item.public,
                                        rolloutStatus: rolloutStatus,
                                        summary: item.summary,
                                        serviceHealthHubState: item.serviceHealthHubState == null ? "PENDING" : item.serviceHealthHubState,
                                        serviceHealthHubViewpoint: item.serviceHealthHubViewpoint,
                                        joinedTags: joinedTags
                                    });
                                }
                            }

                            this._allNews = news;

                            this.setState({
                                items: this._getViewTypeFilter(this._allNews),
                                isDataLoaded: true
                            });
                        });
            }).catch((err) => {
                this.setState({
                    error: err.message,
                    isDataLoaded: true
                });
            });
    }

    private _getViewTypeFilter(source: INewsEvent[]): INewsEvent[] {
        var news: INewsEvent[] = [];
        Object.assign(news, source);

        if (this.props.collectionMode)
            switch (this.props.collectionMode) {
                case NewsItems.Inbox:
                    news = news.filter((n: INewsEvent) => !n.serviceHealthHubViewpoint?.archived);
                    break;
                case NewsItems.Archived:
                    news = news.filter((n: INewsEvent) => n.serviceHealthHubViewpoint?.archived);
                    break;
                case NewsItems.Favorites:
                    news = news.filter((n: INewsEvent) => n.serviceHealthHubViewpoint?.favorite);
                    break;
                default:
                    break;
            }

        return news;
    }

    _onLoadCustomActions(actions: ICustomAction[]): void {
        this.setState({
            customActions: actions
        });
    }

    private _switchViewState(id: string): void {
        var item: INewsEvent | undefined = this._allNews.find(m => m.id.toLowerCase() === id.toLowerCase());

        if (item?.serviceHealthHubViewpoint)
            setViewState(item.id, !item.serviceHealthHubViewpoint.viewed,
                (id: string, state: boolean) => {
                    item!.serviceHealthHubViewpoint.viewed = !item?.serviceHealthHubViewpoint.viewed;
                    this.setState({
                        items: this._getViewTypeFilter(this._allNews)
                    });
                },
                (message: string) => {
                    this.setState({
                        error: message
                    });
                });
    }

    private _switchArchiveState(id: string): void {
        var item: INewsEvent | undefined = this._allNews.find(m => m.id.toLowerCase() === id.toLowerCase());

        if (item?.serviceHealthHubViewpoint)
            setArchiveState(item.id, !item.serviceHealthHubViewpoint.archived,
                (id: string, state: boolean) => {
                    item!.serviceHealthHubViewpoint.archived = !item?.serviceHealthHubViewpoint.archived;
                    this.setState({
                        items: this._getViewTypeFilter(this._allNews)
                    });
                },
                (message: string) => {
                    this.setState({
                        error: message
                    });
                });
    }

    private _switchFavoriteState(id: string): void {
        var item: INewsEvent | undefined = this._allNews.find(m => m.id.toLowerCase() === id.toLowerCase());

        if (item?.serviceHealthHubViewpoint)
            setFavoriteState(item.id, !item.serviceHealthHubViewpoint.favorite,
                (id: string, state: boolean) => {
                    item!.serviceHealthHubViewpoint.favorite = !item?.serviceHealthHubViewpoint.favorite;
                    this.setState({
                        items: this._getViewTypeFilter(this._allNews)
                    });
                },
                (message: string) => {
                    this.setState({
                        error: message
                    });
                });
    }

    _onOpenDetailsPanel(id: string) {
        this.setState({
            selectedItem: id,
            isDetailsPanelOpen: true
        });
    }

    private _onDismisDetailsPanel = (): void => {
        this.setState({
            selectedItem: "",
            isDetailsPanelOpen: false
        });
    }

    private _selectPrevious(id: string): void {
        if (this.state.items !== undefined) {
            const itemIndex: number = this.state.items.findIndex((i) => i.id === id);

            if (itemIndex > 0) {
                this.setState({
                    selectedItem: this.state.items[itemIndex - 1].id
                });
            }
        }
    }

    private _selectNext(id: string): void {
        if (this.state.items !== undefined) {
            const itemIndex: number = this.state.items.findIndex((i) => i.id === id);

            if (itemIndex < this.state.items.length) {
                this.setState({
                    selectedItem: this.state.items[itemIndex + 1].id
                });
            }
        }
    }

    private _isFirst(id: string): boolean {
        if (this.state.items !== undefined && this.state.items.length > 0) {
            const firstElement: INewsEvent = this.state.items[0];
            return firstElement.id === id;
        } else
            return true;
    }

    private _isLast(id: string): boolean {
        if (this.state.items !== undefined && this.state.items.length > 0) {
            const lastElement: INewsEvent = this.state.items[this.state.items.length - 1];
            return lastElement.id === id;
        } else
            return true;
    }
}