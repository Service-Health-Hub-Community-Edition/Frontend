import * as React from 'react';
import { DefaultPalette } from '@fluentui/react';
import { Text, Link, ILinkStyleProps, ILinkStyles, TooltipHost } from '@fluentui/react';
import {
    IColumn, DetailsList, SelectionMode, DetailsListLayoutMode, IDetailsHeaderProps, IRenderFunction, FontIcon, mergeStyles, mergeStyleSets,
    HoverCard, IExpandingCardProps, DirectionalHint, IconButton,
    ScrollablePane, Sticky, ConstrainMode, IContextualMenuProps
} from '@fluentui/react';
import { Panel, PanelType } from '@fluentui/react';
import { Tag, TagList, TagType, ITagProps } from '@m365-admin/tag';
import { AzureUpdateDetails } from './AzureUpdateDetails';
import { ICustomAction, CustomAction } from '../../../CustomAction';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { MessageBar, MessageBarType } from '@fluentui/react';
import { Filter } from "../../../Filter";
import { acquireAccessToken } from "../../../../auth/AccessTokenHelper";
import { checkInTeams } from '../../../auth/detectTeams';
import { AccessDenied } from "../../../AccessDenied";
import { setViewState, setArchiveState, setFavoriteState } from "../../../../api/viewpoint";

export enum AzureUpdatesReportMode {
    dashboard = "dashboard",
    page = "page"
}

export enum AzureUpdatesReportType {
    active,
    history
}

export interface IAzureUpdateEvent {
    id: string;
    title: string;
    summary: string;
    releaseStatus: string;
    contents: string[];
    tags: string[];
    published: Date;
    link: string;
    additionalData: any;
    favorite: boolean;
    serviceHealthHubState: string;
    serviceHealthHubViewpoint: any;
    orgTags: string[];
    joinedTags: string[];
}

interface IAzureUpdatesReportState {
    items: IAzureUpdateEvent[];
    filter?: Filter;
    filterChangeToken?: string;
    isDetailsPanelOpen: boolean;
    selectedMessage: string;
    customActions: ICustomAction[] | undefined;
    isDataLoaded: boolean;
    inTeams: boolean;
    error?: string;
    accessGranted?: boolean;
}

export enum AzureUpdatesItems {
    All = 0,
    Inbox = 1,
    Archived = 2,
    Favorites = 3
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

const componentName: string = 'AzureUpdate';

export class AzureUpdatesReport extends React.Component<{
    type: AzureUpdatesReportType,
    mode?: AzureUpdatesReportMode,
    filter?: Filter,
    filterChangeToken?: string,
    collectionMode?: AzureUpdatesItems,
    onDataLoaded?: (items: IAzureUpdateEvent[]) => void,
    onFilterChange?: (itemCount: number) => void,
    onReloadFilterOptions?: (items: IAzureUpdateEvent[]) => void
}, IAzureUpdatesReportState> {

    private _allItems: IAzureUpdateEvent[] = [];
    customAction: any = React.createRef();

    constructor(props: {
        type: AzureUpdatesReportType,
        mode?: AzureUpdatesReportMode,
        filter?: Filter,
        filterChangeToken?: string,
        collectionMode?: AzureUpdatesItems,
        onDataLoaded?: (items: IAzureUpdateEvent[]) => void,
        onFilterChange?: (itemCount: number) => void,
        onReloadFilterOptions?: (items: IAzureUpdateEvent[]) => void
}) {
        super(props);

        this.state = {
            items: [],
            filter: props.filter,
            filterChangeToken: props.filterChangeToken,
            isDetailsPanelOpen: false,
            selectedMessage: "",
            customActions: undefined,
            isDataLoaded: false,
            inTeams: checkInTeams(),
            error: undefined,
            accessGranted: undefined
        };
    }

    handleView = (itemId: string, viewed: boolean) => {
        var item = this._allItems.find(m => m.id.toLowerCase() === itemId.toLowerCase());
        if (item !== null && item !== undefined && item.serviceHealthHubViewpoint !== undefined) {
            item.serviceHealthHubViewpoint.viewed = viewed;
            this.setState({
                items: this._getFilteredItems(this._allItems)
            });
        }
    }

    handleFavorite = (itemId: string, favorite: boolean) => {
        var item = this._allItems.find(m => m.id.toLowerCase() === itemId.toLowerCase());
        if (item !== null && item !== undefined && item.serviceHealthHubViewpoint !== undefined) {
            item.serviceHealthHubViewpoint.favorited = favorite;
            item.favorite = favorite;
            this.setState({
                items: this._getFilteredItems(this._allItems)
            });
        }
    }

    handleArchive = (itemId: string, archived: boolean) => {
        var item = this._allItems.find(m => m.id.toLowerCase() === itemId.toLowerCase());
        if (item !== null && item !== undefined && item.serviceHealthHubViewpoint !== undefined) {
            item.serviceHealthHubViewpoint.archived = archived;
            this.setState({
                items: this._getViewTypeFilter(this._allItems)
            });
        }
    }

    handleOrgTagsUpdate = (itemId: string, orgTags: string[]) => {
        var item = this._allItems.find(m => m.id.toLowerCase() === itemId.toLowerCase());
        if (item !== null && item !== undefined) {
            item.orgTags = orgTags;

            var viewState = this._getViewTypeFilter(this._allItems);
            var filteredItems: IAzureUpdateEvent[] = this.state.filter ? this.state.filter.filterItems(viewState) : viewState;

            if (this.props.onFilterChange)
                try {
                    this.props.onFilterChange(filteredItems.length);
                }
                catch
                {

                }

            this.setState({
                items: filteredItems
            });



            if (this.props.onReloadFilterOptions)
                try {
                    this.props.onReloadFilterOptions(this._allItems);
                }
                catch
                {

                }
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

        const { items, customActions, isDetailsPanelOpen, selectedMessage, isDataLoaded, error } = this.state;

        const pipeFabricStyles = (p: ILinkStyleProps): ILinkStyles => ({
            root: {
                textDecoration: 'none',
                color: p.theme.semanticColors.bodyText
            },
        });

        const onRenderCompactCard = (item: IAzureUpdateEvent): JSX.Element => {
            return (
                <div className={classNames.compactCard}>
                    <div className="container" style={{ margin: "12px" }} >
                        <div className="row" style={{ paddingBottom: "6px" }} >
                            <div className="col-auto" style={{ marginRight: "0px", paddingRight: "6px" }} >
                                <FontIcon
                                    iconName='InfoSolid'
                                    className={classNames.advisory} />
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
                                    {item.id} · Published {item.published.toLocaleDateString([], { year: 'numeric', month: 'short', day: 'numeric' })}
                                </Text>
                            </div>
                        </div>
                    </div>
                </div>
            );
        };

        const onRenderExpandedCard = (item: IAzureUpdateEvent): JSX.Element => {
            return (
                <div className={classNames.expandedCard}>
                    {!(item.serviceHealthHubState == null || item.serviceHealthHubState == undefined || item.serviceHealthHubState.trim() == "") ? (<>
                        <div style={{ marginBottom: '6px' }}>
                            {renderStatusTag(item.id, item.serviceHealthHubState)}
                        </div></>
                    ) : ""}

                    <>
                        <Text variant='medium'><b>Summary</b></Text><br />
                        <Text variant='small'>{item.summary}</Text>
                    </>
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

        const stateEmpty: boolean = items.find((i: IAzureUpdateEvent) => !(i.serviceHealthHubState == null || i.serviceHealthHubState == undefined || i.serviceHealthHubState == "")) == undefined;

        const columns: IColumn[] = [
            {
                key: 'clType',
                name: '',
                minWidth: 10,
                maxWidth: 10,
                isResizable: false,
                isCollapsible: false,
                isMultiline: false,
                onRender: (item: IAzureUpdateEvent) => {
                    const expandingCardProps: IExpandingCardProps = {
                        onRenderCompactCard: onRenderCompactCard,
                        onRenderExpandedCard: onRenderExpandedCard,
                        renderData: item,
                        directionalHint: DirectionalHint.leftTopEdge,
                    };

                    return <div className="container" style={{ cursor: 'default'}} >
                        <div className="row" >
                            <div className="col" style={{ justifyContent: 'center' }} >
                                <HoverCard expandingCardProps={expandingCardProps}>
                                    <FontIcon
                                        iconName='InfoSolid'
                                        className={classNames.advisory} />
                                </HoverCard>
                            </div>
                        </div></div>;
                },
                isPadded: true,
            },
            {
                key: 'clTitle',
                name: 'Title',
                fieldName: 'title',
                minWidth: 180,
                maxWidth: this.props.mode && this.props.mode === AzureUpdatesReportMode.dashboard ? 350 : 550,
                isResizable: true,
                isCollapsible: false,
                isMultiline: false,
                data: 'string',
                onRender: (item: IAzureUpdateEvent) => {
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

                    root: { padding: '6px 0px 6px 0px' }
                },
                onRender: (item: IAzureUpdateEvent) => {
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
                onRender: (item: IAzureUpdateEvent) => {
                    return renderStatusTag(item.id, item.serviceHealthHubState)
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
                onRender: (item: IAzureUpdateEvent) => {
                    if (this.props.mode === AzureUpdatesReportMode.page)
                        return item.tags ? (<TagList tags={item.tags.map((tag: string) => ({ key: tag, children: tag, }))} />) : ""
                    else
                        return item.joinedTags ? (<TagList tags={item.joinedTags.map((tag: string) => ({ key: tag, children: tag, }))} />) : ""
                },
            }]

        if (this.props.mode === AzureUpdatesReportMode.page)
            columns.push({
                key: 'clOrgTags',
                name: 'Org. tags',
                fieldName: 'orgTags',
                minWidth: 80,
                maxWidth: 300,
                isResizable: true,
                isCollapsible: true,
                data: 'string',
                isPadded: true,
                onRender: (item: IAzureUpdateEvent) => {
                    return item.orgTags ? (<TagList tags={item.orgTags.map((tag: string) => ({ key: tag, children: tag, }))} />) : ""
                },
            });

        columns.push({
            key: 'clUpdated',
            name: 'Updated',
            minWidth: 120,
            maxWidth: 120,
            isResizable: false,
            isCollapsible: true,
            isPadded: false,
            onRender: (item: IAzureUpdateEvent) => {
                return <>{item.published ?
                    new Intl.DateTimeFormat(
                        window.navigator.language,
                        {
                            day: 'numeric',
                            month: 'numeric',
                            year: 'numeric',
                            hour: 'numeric',
                            minute: 'numeric'
                        }).format(item.published) : ""}</>;
            },
        });
        
        columns.push({
            key: 'clId',
            name: 'Id',
            minWidth: 80,
            maxWidth: 80,
            isResizable: true,
            isCollapsible: true,
            data: 'string',
            isPadded: false,
            onRender: (item: IAzureUpdateEvent) => {
                return <>{item.id}</>;
            },
        });

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

            { this.props.mode !== AzureUpdatesReportMode.dashboard ? (
                <DetailsList
                    items={items}
                    compact={true}
                    columns={columns}
                    selectionMode={SelectionMode.none}
                    layoutMode={DetailsListLayoutMode.justified}
                    constrainMode={ConstrainMode.horizontalConstrained}
                    isHeaderVisible={true}
                />) : (
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
            </ScrollablePane>)}

            <Panel
                    headerText={items.find(m => m.id.toLowerCase() === selectedMessage.toLowerCase())?.title!}
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
                                disabled={this._isFirst(selectedMessage)}
                                onClick={() => this._selectPrevious(selectedMessage)} />

                            <IconButton
                                iconProps={{ iconName: 'Down' }}
                                title='Next'
                                ariaLabel='Next item'
                                disabled={this._isLast(selectedMessage)}
                                onClick={() => this._selectNext(selectedMessage)} />

                            <IconButton
                                iconProps={{ iconName: 'Cancel' }}
                                title='Close'
                                ariaLabel='Close panel'
                                onClick={() => this._onDismisDetailsPanel()} />

                            {defaultRender!(props)}
                        </div>
                    )}
                    >
                <AzureUpdateDetails id={selectedMessage} onView={this.handleView} onUpdateOrgTags={this.handleOrgTagsUpdate} />
            </Panel>
        </>
        );
    }

    componentDidMount() {
        this._getAzureUpdates();
    }

    componentDidUpdate(prevProps: any) {

        if (this.props.filterChangeToken && this.props.filterChangeToken !== this.state.filterChangeToken) {
            var viewState = this._getViewTypeFilter(this._allItems);
            var filteredItems: IAzureUpdateEvent[] = this.state.filter ? this.state.filter.filterItems(viewState) : viewState;

            if (this.props.onFilterChange)
                try {
                    this.props.onFilterChange(filteredItems.length);
                }
                catch
                {

                }

            this.setState({
                filterChangeToken: this.props.filterChangeToken,
                items: filteredItems
            });
        }
    }

    private _getViewTypeFilter(source: IAzureUpdateEvent[]): IAzureUpdateEvent[] {
        var events: IAzureUpdateEvent[] = [];
        Object.assign(events, source);

        if (this.props.collectionMode)
            switch (this.props.collectionMode) {
                case AzureUpdatesItems.Inbox:
                    events = events.filter((e: IAzureUpdateEvent) => !e.serviceHealthHubViewpoint?.archived);
                    break;
                case AzureUpdatesItems.Archived:
                    events = events.filter((e: IAzureUpdateEvent) => e.serviceHealthHubViewpoint?.archived);
                    break;
                case AzureUpdatesItems.Favorites:
                    events = events.filter((e: IAzureUpdateEvent) => e.serviceHealthHubViewpoint?.favorite);
                    break;
                default:
                    break;
            }

        return events;
    }

    _onLoadCustomActions(actions: ICustomAction[]): void {
        this.setState({
            customActions: actions
        });
    }

    private _getFilteredItems(items: IAzureUpdateEvent[]): IAzureUpdateEvent[] {
        var viewState = this._getViewTypeFilter(this._allItems);
        var filteredItems: IAzureUpdateEvent[] = this.state.filter ? this.state.filter.filterItems(viewState) : viewState;

        if (this.props.onFilterChange)
            this.props.onFilterChange(filteredItems.length);

        return filteredItems;
    }

    private _switchViewState(id: string): void {
        var item: IAzureUpdateEvent | undefined = this._allItems.find(m => m.id.toLowerCase() === id.toLowerCase());

        if (item?.serviceHealthHubViewpoint)
            setViewState(item.id, !item.serviceHealthHubViewpoint.viewed,
                (id: string, state: boolean) => {
                    item!.serviceHealthHubViewpoint.viewed = !item?.serviceHealthHubViewpoint.viewed;
                    this.setState({
                        items: this._getFilteredItems(this._allItems)
                    });
                },
                (message: string) => {
                    this.setState({
                        error: message
                    });
                });
    }

    private _switchArchiveState(id: string): void {
        var item: IAzureUpdateEvent | undefined = this._allItems.find(m => m.id.toLowerCase() === id.toLowerCase());

        if (item?.serviceHealthHubViewpoint)
            setArchiveState(item.id, !item.serviceHealthHubViewpoint.archived,
                (id: string, state: boolean) => {
                    item!.serviceHealthHubViewpoint.archived = !item?.serviceHealthHubViewpoint.archived;
                    this.setState({
                        items: this._getFilteredItems(this._allItems)
                    });
                },
                (message: string) => {
                    this.setState({
                        error: message
                    });
                });
    }

    private _switchFavoriteState(id: string): void {
        var item: IAzureUpdateEvent | undefined = this._allItems.find(m => m.id.toLowerCase() === id.toLowerCase());

        if (item?.serviceHealthHubViewpoint)
            setFavoriteState(item.id, !item.serviceHealthHubViewpoint.favorite,
                (id: string, state: boolean) => {
                    item!.serviceHealthHubViewpoint.favorited = !item?.serviceHealthHubViewpoint.favorited;
                    item!.favorite = item!.serviceHealthHubViewpoint.favorited;
                    this.setState({
                        items: this._getFilteredItems(this._allItems)
                    });
                },
                (message: string) => {
                    this.setState({
                        error: message
                    });
                });
    }

    private _getAzureUpdates() {
        var updates: IAzureUpdateEvent[] = [];
        const requiredRoles: string[] = ['ServiceHealthReader', 'Communication.Write.All', 'Admin'];

        acquireAccessToken()
            .then((response) => {
                var tokenClaims: any = response.account?.idTokenClaims;
                const userRoles: any = tokenClaims?.roles;
                var userHasRequiredRole: boolean = userRoles.some((r: string) => requiredRoles.includes(r));

                this.setState({
                    accessGranted: userHasRequiredRole
                });

                let url: string = '/api/azure/updates';

                if (this.props.type !== undefined) {
                    let query: string = "";

                    switch (this.props.type) {
                        case AzureUpdatesReportType.active:
                            query = "?active=true";
                            break;
                        case AzureUpdatesReportType.history:
                            query = "?active=false";
                            break;
                        default:
                            query = ""
                    }

                    url = url + query;
                }
                
                if (userHasRequiredRole)
                    fetch(url, { headers: response.idToken === "" ? {} : { 'Authorization': `Bearer ${response.idToken}` } })
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
                                var sortedResult = result.sort((a: any, b: any) => new Date(a.lastUpdatedTime) > new Date(b.lastUpdatedTime) ? -1 : 1);
                                for (const update of sortedResult) {
                                    var data: any = JSON.parse(update.data);

                                    var orgTags: string[] = update.organizationTags.map((ot: any) => ot.tag);
                                    var tags: string[] = Array.isArray(data.tags) ? data.tags : (data.tags === null ? [] : [data.tags]);
                                    var joinedTags: string[] = [];

                                    for (const t of tags)
                                        joinedTags.push(t);

                                    for (const t of orgTags)
                                        joinedTags.push('[ORG] ' + t);

                                    updates.push({
                                        id: update.id,
                                        title: data.title,
                                        summary: data.summary,
                                        published: new Date(data.published),
                                        contents: data.contents,
                                        releaseStatus: Array.isArray(data.releaseStatus) ? data.releaseStatus.join(", ") : data.releaseStatus,
                                        link: data.link,
                                        tags: tags,
                                        additionalData: update.additionalData,
                                        favorite: update.serviceHealthHubViewpoint?.favorited ? update.serviceHealthHubViewpoint.favorited : false,
                                        serviceHealthHubState: update.state,
                                        serviceHealthHubViewpoint: update.serviceHealthHubViewpoint,
                                        orgTags: orgTags,
                                        joinedTags: joinedTags
                                    });
                                }
                            }

                            try {
                                if (this.props.onDataLoaded)
                                    this.props.onDataLoaded(updates)
                            }
                            catch (err) {
                                this.setState({
                                    error: err.message
                                });
                            }

                            this._allItems = updates;
                            var filteredItems: IAzureUpdateEvent[] = this.state.filter ? this.state.filter.filterItems(this._allItems) : this._allItems;

                            if (this.props.onFilterChange)
                                try {
                                    this.props.onFilterChange(filteredItems.length);
                                }
                                catch
                                {

                                }

                            this.setState({
                                items: filteredItems,
                                isDataLoaded: true
                            });
                        });
            }).catch((err) => {
                this.setState({
                    error: err.message
                });
            });
    }


    _onOpenDetailsPanel(id: string) {
        this.setState({
            selectedMessage: id,
            isDetailsPanelOpen: true
        });
    }

    private _onDismisDetailsPanel = (): void => {
        this.setState({
            selectedMessage: "",
            isDetailsPanelOpen: false
        });
    }

    private _selectPrevious(id: string): void {
        if (this.state.items !== undefined) {
            const itemIndex: number = this.state.items.findIndex((i) => i.id === id);

            if (itemIndex > 0) {
                this.setState({
                    selectedMessage: this.state.items[itemIndex - 1].id
                });
            }
        }
    }

    private _selectNext(id: string): void {
        if (this.state.items !== undefined) {
            const itemIndex: number = this.state.items.findIndex((i) => i.id === id);

            if (itemIndex < this.state.items.length) {
                this.setState({
                    selectedMessage: this.state.items[itemIndex + 1].id
                });
            }
        }
    }

    private _isFirst(id: string): boolean {
        if (this.state.items !== undefined && this.state.items.length > 0) {
            const firstElement: IAzureUpdateEvent = this.state.items[0];
            return firstElement.id === id;
        } else
            return true;
    }

    private _isLast(id: string): boolean {
        if (this.state.items !== undefined && this.state.items.length > 0) {
            const lastElement: IAzureUpdateEvent = this.state.items[this.state.items.length - 1];
            return lastElement.id === id;
        } else
            return true;
    }
}