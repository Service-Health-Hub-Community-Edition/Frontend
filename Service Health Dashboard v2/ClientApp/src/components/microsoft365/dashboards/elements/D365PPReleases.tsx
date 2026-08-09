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
import { D365PPReleaseDetails } from './D365PPReleaseDetails';
import { ICustomAction, CustomAction } from '../../../CustomAction';
import { ServiceComponent } from '../../../ServiceNameComponent';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { MessageBar, MessageBarType } from '@fluentui/react';
import { Filter } from "../../../Filter";
import { acquireAccessToken } from "../../../../auth/AccessTokenHelper";
import { checkInTeams } from '../../../auth/detectTeams';
import { AccessDenied } from "../../../AccessDenied";
import { setViewState, setArchiveState, setFavoriteState } from "../../../../api/viewpoint";

export enum D365PPReportMode {
    dashboard = "dashboard",
    page = "page"
}

export interface ID365PPRelease {
    id: string;
    title: string;
    businessValue: string;
    description: string;
    featureType: string;
    product: string;
    productArea: string;
    parentProduct: string;
    enabledFor: string;
    releaseWaveId?: string;
    releaseWave?: string;
    rwStartShipDate?: Date;
    rwEndShipDate?: Date;
    rwStatus?: string;
    earlyAccessDate?: Date;
    earlyAccessStatus?: string;
    publicPreviewDate?: Date;
    publicPreviewStatus?: string;
    gaDate?: Date;
    gaStatus?: string;
    documentation?: string;
    blogArticle?: string;
    overviewVideo?: string;
    published?: Date;
    lastUpdate?: Date;
    favorite: boolean;
    serviceHealthHubState: string;
    serviceHealthHubViewpoint: any;
    shhImageMetadata?: any;
    workItemId?: string;
    workItemUrl?: string;
    extendedProperties?: any;
    orgTags: string[];
}

interface ID365PPReportState {
    items: ID365PPRelease[];
    filter?: Filter;
    filterChangeToken?: string;
    isDetailsPanelOpen: boolean;
    selectedMessage: string;
    selectedMessageTitle: string;
    disablePanelNavigation: boolean;
    customActions: ICustomAction[] | undefined;
    isDataLoaded: boolean;
    inTeams: boolean;
    error?: string;
    accessGranted?: boolean;
}

export enum D365PPReportItems {
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

const componentName: string = 'D365PowerPlatformRelease';

export class D365PPReport extends React.Component<{
    mode?: D365PPReportMode,
    filter?: Filter,
    filterChangeToken?: string,
    collectionMode?: D365PPReportItems,
    onDataLoaded?: (items: ID365PPRelease[]) => void,
    onFilterChange?: (itemCount: number) => void,
    onReloadFilterOptions?: (items: ID365PPRelease[]) => void
}, ID365PPReportState> {

    private _allItems: ID365PPRelease[] = [];
    customAction: any = React.createRef();

    constructor(props: {
        mode?: D365PPReportMode,
        filter?: Filter,
        filterChangeToken?: string,
        collectionMode?: D365PPReportItems,
        onDataLoaded?: (items: ID365PPRelease[]) => void,
        onFilterChange?: (itemCount: number) => void,
        onReloadFilterOptions?: (items: ID365PPRelease[]) => void
}) {
        super(props);

        const queryParams = new URLSearchParams(window.location.search);
        const id = queryParams.get('id');

        this.state = {
            items: [],
            filter: props.filter,
            filterChangeToken: props.filterChangeToken,
            isDetailsPanelOpen: !(id === null || id === undefined),
            selectedMessage: id === null || id === undefined ? "" : id!,
            selectedMessageTitle: "",
            disablePanelNavigation: !(id === null || id === undefined),
            customActions: undefined,
            isDataLoaded: false,
            inTeams: checkInTeams(),
            error: undefined,
            accessGranted: undefined
        };

        this._updatePanelTitle = this._updatePanelTitle.bind(this);
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
            var filteredItems: ID365PPRelease[] = this.state.filter ? this.state.filter.filterItems(viewState) : viewState;

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

        const { items, customActions, isDetailsPanelOpen, selectedMessage, selectedMessageTitle, isDataLoaded, disablePanelNavigation, error } = this.state;

        const pipeFabricStyles = (p: ILinkStyleProps): ILinkStyles => ({
            root: {
                textDecoration: 'none',
                color: p.theme.semanticColors.bodyText
            },
        });

        const onRenderCompactCard = (item: ID365PPRelease): JSX.Element => {
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
                                    {item.published ? (<>Published {item.published.toLocaleDateString([], { year: 'numeric', month: 'short', day: 'numeric' })} </>) : (<></>)}&nbsp;·&nbsp;
                                    {item.lastUpdate ? (<>Updated {item.lastUpdate.toLocaleDateString([], { year: 'numeric', month: 'short', day: 'numeric' })} </>) : (<></>)}
                                </Text>
                            </div>
                        </div>
                    </div>
                </div>
            );
        };

        const onRenderExpandedCard = (item: ID365PPRelease): JSX.Element => {
            return (
                <div className={classNames.expandedCard}>
                    {!(item.serviceHealthHubState == null || item.serviceHealthHubState == undefined || item.serviceHealthHubState.trim() == "") ? (<>
                        <div style={{ marginBottom: '6px' }}>
                            {renderStatusTag(item.id, item.serviceHealthHubState)}
                        </div></>
                    ) : ""}

                    <>
                        <Text variant='medium'><b>Business value</b></Text><br />
                        <Text variant='small'>
                            <div style={{ whiteSpace: "pre-wrap" }} dangerouslySetInnerHTML={{ __html: item.businessValue }} />
                        </Text>
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

        const stateEmpty: boolean = items.find((i: ID365PPRelease) => !(i.serviceHealthHubState == null || i.serviceHealthHubState == undefined || i.serviceHealthHubState == "")) == undefined;

        const columns: IColumn[] = [
            {
                key: 'clType',
                name: '',
                minWidth: 10,
                maxWidth: 10,
                isResizable: false,
                isCollapsible: false,
                isMultiline: false,
                onRender: (item: ID365PPRelease) => {
                    const expandingCardProps: IExpandingCardProps = {
                        onRenderCompactCard: onRenderCompactCard,
                        onRenderExpandedCard: onRenderExpandedCard,
                        renderData: item,
                        directionalHint: DirectionalHint.leftTopEdge,
                    };

                    return <div className="container" style={{ cursor: 'default' }} >
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
                maxWidth: this.props.mode && this.props.mode === D365PPReportMode.dashboard ? 350 : 550,
                isResizable: true,
                isCollapsible: false,
                isMultiline: false,
                data: 'string',
                onRender: (item: ID365PPRelease) => {
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
                onRender: (item: ID365PPRelease) => {
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
                onRender: (item: ID365PPRelease) => {
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
                onRender: (item: ID365PPRelease) => {
                    return item.orgTags ? (<TagList tags={item.orgTags.map((tag: string) => ({ key: tag, children: tag, }))} />) : ""
                }
            },
            {
                key: 'clProduct',
                name: 'Product',
                minWidth: 120,
                maxWidth: 120,
                isResizable: false,
                isCollapsible: true,
                isPadded: false,
                onRender: (item: ID365PPRelease) => {
                    return <>
                        <TooltipHost content={item.product}>
                            <span>
                                <ServiceComponent name={item.product} />
                            </span>
                        </TooltipHost>
                    </>;
                }
            }
        ];

        if (this.props.mode === D365PPReportMode.page)
            columns.push({
                key: 'clProductArea',
                name: 'Product area',
                minWidth: 120,
                maxWidth: 120,
                isResizable: false,
                isCollapsible: true,
                isPadded: false,
                onRender: (item: ID365PPRelease) => {
                    return <>{item.productArea}</>;
                }
            });

        columns.push({
            key: 'clUpdated',
            name: 'Updated',
            minWidth: 120,
            maxWidth: 120,
            isResizable: false,
            isCollapsible: true,
            isPadded: false,
            onRender: (item: ID365PPRelease) => {
                return <>{item.lastUpdate ?
                    new Intl.DateTimeFormat(
                        window.navigator.language,
                        {
                            day: 'numeric',
                            month: 'numeric',
                            year: 'numeric',
                            hour: 'numeric',
                            minute: 'numeric'
                        }).format(item.lastUpdate) : ""}</>;
            }
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

            { this.props.mode !== D365PPReportMode.dashboard ? (
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
                    headerText={selectedMessageTitle}
                    isOpen={isDetailsPanelOpen}
                    onDismiss={this._onDismisDetailsPanel}
                    type={PanelType.medium}
                    // You MUST provide this prop! Otherwise screen readers will just say "button" with no label.
                    closeButtonAriaLabel="Close"
                    hasCloseButton={false}
                    onRenderNavigationContent={(props, defaultRender) => (
                        <div>
                            {disablePanelNavigation ? (<></>) : (<>
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
                            </>)};

                            <IconButton
                                iconProps={{ iconName: 'Cancel' }}
                                title='Close'
                                ariaLabel='Close panel'
                                onClick={() => this._onDismisDetailsPanel()} />

                            {defaultRender!(props)}
                        </div>
                    )}
                    >
                <D365PPReleaseDetails
                    id={selectedMessage}
                    onView={this.handleView}
                    onUpdateOrgTags={this.handleOrgTagsUpdate}
                    onUpdateParentTitle={this._updatePanelTitle}
                />
            </Panel>
        </>
        );
    }

    componentDidMount() {
        this._getD365PP();
    }

    componentDidUpdate(prevProps: any) {

        if (this.props.filterChangeToken && this.props.filterChangeToken !== this.state.filterChangeToken) {
            var viewState = this._getViewTypeFilter(this._allItems);
            var filteredItems: ID365PPRelease[] = this.state.filter ? this.state.filter.filterItems(viewState) : viewState;

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

    private _getViewTypeFilter(source: ID365PPRelease[]): ID365PPRelease[] {
        var events: ID365PPRelease[] = [];
        Object.assign(events, source);

        if (this.props.collectionMode)
            switch (this.props.collectionMode) {
                case D365PPReportItems.Inbox:
                    events = events.filter((e: ID365PPRelease) => !e.serviceHealthHubViewpoint?.archived);
                    break;
                case D365PPReportItems.Archived:
                    events = events.filter((e: ID365PPRelease) => e.serviceHealthHubViewpoint?.archived);
                    break;
                case D365PPReportItems.Favorites:
                    events = events.filter((e: ID365PPRelease) => e.serviceHealthHubViewpoint?.favorite);
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

    private _getFilteredItems(items: ID365PPRelease[]): ID365PPRelease[] {
        var viewState = this._getViewTypeFilter(this._allItems);
        var filteredItems: ID365PPRelease[] = this.state.filter ? this.state.filter.filterItems(viewState) : viewState;

        if (this.props.onFilterChange)
            this.props.onFilterChange(filteredItems.length);

        return filteredItems;
    }

    private _switchViewState(id: string): void {
        var item: ID365PPRelease | undefined = this._allItems.find(m => m.id.toLowerCase() === id.toLowerCase());

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
        var item: ID365PPRelease | undefined = this._allItems.find(m => m.id.toLowerCase() === id.toLowerCase());

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
        var item: ID365PPRelease | undefined = this._allItems.find(m => m.id.toLowerCase() === id.toLowerCase());

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

    private _getD365PP() {
        var updates: ID365PPRelease[] = [];
        const requiredRoles: string[] = ['ServiceHealthReader', 'Communication.Write.All', 'Admin'];

        acquireAccessToken()
            .then((response) => {
                var tokenClaims: any = response.account?.idTokenClaims;
                const userRoles: any = tokenClaims?.roles;
                var userHasRequiredRole: boolean = userRoles.some((r: string) => requiredRoles.includes(r));

                this.setState({
                    accessGranted: userHasRequiredRole
                });

                let url: string = '/api/d365pp/releases';
                
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
                                var sortedResult = result.sort((a: any, b: any) => new Date(a.lastUpdate) > new Date(b.lastUpdate) ? -1 : 1);

                                for (const update of sortedResult) {
                                    var orgTags: string[] = update.organizationTags.map((ot: any) => ot.tag);

                                    updates.push({
                                        id: update.id,
                                        title: update.title,
                                        businessValue: update.businessValue,
                                        description: update.description,
                                        featureType: update.featureType,
                                        product: update.product,
                                        productArea: update.productArea,
                                        parentProduct: update.parentProduct,
                                        enabledFor: update.enabledFor,
                                        releaseWaveId: update.releaseWaveId,
                                        releaseWave: update.releaseWave,
                                        rwStartShipDate: update.rwStartShipDate,
                                        rwEndShipDate: update.rwEndShipDate,
                                        rwStatus: update.rW_Status,
                                        earlyAccessDate: update.earlyAccessDate,
                                        earlyAccessStatus: update.earlyAccessStatus,
                                        publicPreviewDate: update.publicPreviewDate,
                                        publicPreviewStatus: update.publicPreviewStatus,
                                        gaDate: update.gaDate,
                                        gaStatus: update.gaStatus,
                                        documentation: update.documentation,
                                        blogArticle: update.blogArticle,
                                        overviewVideo: update.overviewVideo,
                                        lastUpdate: update.lastUpdate ? new Date(update.lastUpdate) : undefined,
                                        published: update.published ? new Date(update.published) : undefined,
                                        shhImageMetadata: update.shhImageMetadata,
                                        favorite: update.serviceHealthHubViewpoint?.favorited ? update.serviceHealthHubViewpoint.favorited : false,
                                        serviceHealthHubState: update.state ? update.state : "",
                                        serviceHealthHubViewpoint: update.serviceHealthHubViewpoint,
                                        orgTags: orgTags
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
                            var filteredItems: ID365PPRelease[] = this.state.filter ? this.state.filter.filterItems(this._allItems) : this._allItems;

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

    _updatePanelTitle(title: string) {
        this.setState({
            selectedMessageTitle: title
        });
    }
    
    _onOpenDetailsPanel(id: string) {
        this.setState({
            selectedMessage: id,
            selectedMessageTitle: this.state.items.find(m => m.id.toLowerCase() === id.toLowerCase())?.title!,
            disablePanelNavigation: false,
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
                    selectedMessage: this.state.items[itemIndex - 1].id,
                    selectedMessageTitle: this.state.items.find(m => m.id.toLowerCase() === this.state.items[itemIndex - 1].id.toLowerCase())?.title!,
                    disablePanelNavigation: false
                });
            }
        }
    }

    private _selectNext(id: string): void {
        if (this.state.items !== undefined) {
            const itemIndex: number = this.state.items.findIndex((i) => i.id === id);

            if (itemIndex < this.state.items.length) {
                this.setState({
                    selectedMessage: this.state.items[itemIndex + 1].id,
                    selectedMessageTitle: this.state.items.find(m => m.id.toLowerCase() === this.state.items[itemIndex + 1].id.toLowerCase())?.title!,
                    disablePanelNavigation: false
                });
            }
        }
    }

    private _isFirst(id: string): boolean {
        if (this.state.items !== undefined && this.state.items.length > 0) {
            const firstElement: ID365PPRelease = this.state.items[0];
            return firstElement.id === id;
        } else
            return true;
    }

    private _isLast(id: string): boolean {
        if (this.state.items !== undefined && this.state.items.length > 0) {
            const lastElement: ID365PPRelease = this.state.items[this.state.items.length - 1];
            return lastElement.id === id;
        } else
            return true;
    }
}