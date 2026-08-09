import * as React from 'react';
import { ColumnActionsMode, DefaultPalette } from '@fluentui/react';
import { Text, Link, ILinkStyleProps, ILinkStyles, TooltipHost, ScrollablePane, Sticky, IDetailsHeaderProps, IRenderFunction } from '@fluentui/react';
import {
    IColumn, DetailsList, SelectionMode, DetailsListLayoutMode, ConstrainMode, FontIcon,
    IDetailsListStyles, mergeStyles, mergeStyleSets, IContextualMenuItem,
    HoverCard, IExpandingCardProps, DirectionalHint, IconButton, IContextualMenuProps
} from '@fluentui/react';
import { Panel, PanelType } from '@fluentui/react';
import { Tag, TagList, TagType, ITagProps } from '@m365-admin/tag';
import { IncidentDetails } from '../../../IncidentDetails';
import { ICustomAction, CustomAction } from '../../../CustomAction';
import { ServiceComponent } from '../../../ServiceNameComponent';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { MessageBar, MessageBarType } from '@fluentui/react';
import { acquireAccessToken, accessControl } from "../../../../auth/AccessTokenHelper";
import { checkInTeams } from '../../../auth/detectTeams';
import { AccessDenied } from "../../../AccessDenied";
import { Filter } from "../../../Filter";
import { setViewState, setArchiveState, setFavoriteState } from "../../../../api/viewpoint";

export interface ISummaryElement {
    text: string;
}

export interface ISummary {
    timestamp?: Date;
    contents?: ISummaryElement[];
}

export interface IHealthEvent {
    id: string;
    title: string;
    service: string[];
    issueType: string;
    status: string;
    start: Date;
    lastModified: Date;
    isResolved: boolean; 
    published?: boolean;
    favorite?: boolean;
    origin: string;
    summary?: ISummary;
    summarySearchable?: string;
    serviceHealthHubState: string;
    serviceHealthHubViewpoint: any;
    orgTags: string[];
}

interface IHealthReportState {
    filter?: Filter;
    filterChangeToken?: string;
    items: IHealthEvent[];
    customActions: ICustomAction[] | undefined;
    isDetailsPanelOpen: boolean;
    selectedIncident: string;
    selectedIncidentTitle: string;
    disablePanelNavigation: boolean;
    isDataLoaded: boolean;
    inTeams: boolean;
    error?: string;
    accessGranted?: boolean;
}

export enum HealthReportType {
    active,
    history
}

export enum HealthItems {
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

export enum HealthReportMode {
    dashboard = "dashboard",
    page = "page"
}

const componentName: string = 'ServiceHealthIssue';

export class HealthReport extends React.Component<{
    type: HealthReportType,
    mode?: HealthReportMode,
    collectionMode?: HealthItems,
    filter?: Filter,
    filterChangeToken?: string,
    onDataLoaded?: (items: IHealthEvent[]) => void,
    onFilterChange?: (itemCount: number) => void,
    onReloadFilterOptions?: (items: IHealthEvent[]) => void
}, IHealthReportState> {
    private _allItems: IHealthEvent[] = [];
    customAction: any = React.createRef();

    constructor(props: {
        type: HealthReportType,
        mode?: HealthReportMode,
        collectionMode?: HealthItems,
        filter?: Filter,
        filterChangeToken?: string,
        onDataLoaded?: (items: IHealthEvent[]) => void,
        onFilterChange?: (itemCount: number) => void,
        onReloadFilterOptions?: (items: IHealthEvent[]) => void
    }) {
        super(props);

        const queryParams = new URLSearchParams(window.location.search);
        const id = queryParams.get('id');

        this.state = {
            filter: props.filter,
            filterChangeToken: props.filterChangeToken,
            items: [],
            customActions: undefined,
            isDetailsPanelOpen: !(id === null || id === undefined),
            selectedIncident: id === null || id === undefined ? "" : id!,
            selectedIncidentTitle: "",
            disablePanelNavigation: !(id === null || id === undefined),
            isDataLoaded: false,
            inTeams: checkInTeams(),
            error: undefined,
            accessGranted: undefined
        };

        this._updatePanelTitle = this._updatePanelTitle.bind(this);
    }

    handleItemChange = (itemId: string, published: boolean) => {
        var item = this._allItems.find(m => m.id.toLowerCase() === itemId.toLowerCase());
        if (item !== null && item !== undefined) {
            item.published = published;
            this.setState({
                items: this._getFilteredItems(this._allItems)
            });
        }
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
            var filteredItems: IHealthEvent[] = this.state.filter ? this.state.filter.filterItems(viewState) : viewState;

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
                    console.log("Entering onReloadFilterOptions");
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

        const { items, customActions, isDetailsPanelOpen, selectedIncident, selectedIncidentTitle, disablePanelNavigation, isDataLoaded, error } = this.state;

        const pipeFabricStyles = (p: ILinkStyleProps): ILinkStyles => ({
            root: {
                textDecoration: 'none',
                color: p.theme.semanticColors.bodyText
            },
        });

        const onRenderCompactCard = (item: IHealthEvent): JSX.Element => {
            return (
                <div className={classNames.compactCard}>
                    <div className="container" style={{ margin: "12px" }} >
                        <div className="row" style={{ paddingBottom: "6px" }} >
                            <div className="col-auto" style={{ marginRight: "0px", paddingRight: "6px" }} >
                                <FontIcon
                                    iconName={item.issueType === 'Incident' ? 'WarningSolid' : 'InfoSolid'}
                                    className={item.issueType === 'Incident' ? classNames.incident : classNames.advisory} />
                            </div>
                            <div className="col" style={{ margin: "0px", padding: "0px", alignItems: 'center' }}>
                                <Text variant='smallPlus'><b>{item.issueType}</b></Text>
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
                                    {item.id} · Start {item.start.toLocaleDateString([], { year: 'numeric', month: 'short', day: 'numeric' })} · Updated {item.lastModified.toLocaleDateString([], { year: 'numeric', month: 'short', day: 'numeric' })}
                                </Text>
                            </div>
                        </div>
                    </div>
                </div>
            );
        };

        const onRenderExpandedCard = (item: IHealthEvent): JSX.Element => {
            return (
                <div className={classNames.expandedCard}>
                    {item.summary?.contents && item.summary?.contents.length > 0 ? (
                        <>
                            {item.published || !(item.serviceHealthHubState == null || item.serviceHealthHubState == undefined || item.serviceHealthHubState.trim() == "") ? (<>
                                <div style={{ marginBottom: '6px' }}>
                                    {renderStatusTag(item.id, item.serviceHealthHubState)}
                                    &nbsp;
                                    {item.published ? (<Tag
                                        onRenderContent={(tagProps: ITagProps | undefined, defaultRender: IRenderFunction<ITagProps> | undefined): JSX.Element => {
                                            return <><FontIcon iconName='Streaming' /> {defaultRender && tagProps ? defaultRender(tagProps) : ""}</>;
                                        }} >
                                        Published
                                    </Tag>) : (<></>)} </div></>
                            ) : ""}
                            <Text variant='medium'><b>Summary</b></Text><br />
                            {item.summary.timestamp ? (
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
                        <div style={{ textAlign: 'center' }}>
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

        const stateEmpty: boolean = items.find((i: IHealthEvent) => !(i.serviceHealthHubState == null || i.serviceHealthHubState == undefined || i.serviceHealthHubState == "")) == undefined;

        var columns: IColumn[] = [
            {
                key: 'clType',
                name: '',
                minWidth: 4,
                maxWidth: 4,
                isResizable: false,
                isCollapsible: false,
                isMultiline: false,
                currentWidth: 4,
                onRender: (item: IHealthEvent) => {
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
                                    {item.published ? (<FontIcon iconName='Streaming' />) : ""}
                                </div>
                                <div className="col" style={{ justifyContent: 'center', paddingLeft: '0px', paddingRight: '3px' }} >

                                    <FontIcon
                                        iconName={item.issueType === 'Incident' ? 'WarningSolid' : 'InfoSolid'}
                                        className={item.issueType === 'Incident' ? classNames.incident : classNames.advisory} />

                                </div>
                            </div></HoverCard></div>;
                },
                isPadded: true,
            },
            {
                key: 'clTitle',
                name: 'Title',
                fieldName: 'title',
                minWidth: 180,
                maxWidth: this.props.mode && this.props.mode === HealthReportMode.dashboard ? 350 : 550,
                isResizable: true,
                isCollapsible: false,
                isMultiline: false,
                data: 'string',
                onRender: (item: IHealthEvent) => {
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
                onRender: (item: IHealthEvent) => {
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
                onRender: (item: IHealthEvent) => {
                    return renderStatusTag(item.id, item.serviceHealthHubState)
                },
            },
            {
                key: 'clServices',
                name: 'Affected services',
                minWidth: 100,
                maxWidth: this.props.mode && this.props.mode === HealthReportMode.dashboard ? 100 : 150,
                isResizable: this.props.mode && this.props.mode === HealthReportMode.dashboard ? false : true,
                isCollapsible: true,
                isPadded: false,
                onRender: (item: IHealthEvent) => {
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
                key: 'clStatus',
                name: 'Status',
                fieldName: 'status',
                minWidth: 120,
                maxWidth: 120,
                isResizable: this.props.mode && this.props.mode === HealthReportMode.dashboard ? false : true,
                isCollapsible: true,
                data: 'string',
                isPadded: false,
                onRender: (item: IHealthEvent) => {
                    return <>
                        <TooltipHost
                            content={item.status} >
                            {item.status}
                        </TooltipHost>
                    </>;
                },
            },
            {
                key: 'clUpdated',
                name: 'Updated',
                minWidth: 120,
                maxWidth: 120,
                isResizable: false,
                isCollapsible: true,
                isPadded: false,
                onRender: (item: IHealthEvent) => {
                    return <>{item.lastModified ?
                        new Intl.DateTimeFormat(
                            window.navigator.language,
                            {
                                day: 'numeric',
                                month: 'numeric',
                                year: 'numeric',
                                hour: 'numeric',
                                minute: 'numeric'
                            }).format(item.lastModified) : ""}</>;
                },
            }];

        if (this.props.mode === HealthReportMode.page)
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
                onRender: (item: IHealthEvent) => {
                    return item.orgTags ? (<TagList tags={item.orgTags.map((tag: string) => ({ key: tag, children: tag, }))} />) : ""
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
                onRender: (item: IHealthEvent) => {
                    return <>{item.id}</>;
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

            {this.props.mode !== HealthReportMode.dashboard ? (
                <DetailsList
                    items={items}
                    compact={true}
                    columns={columns}
                    selectionMode={SelectionMode.none}
                    layoutMode={DetailsListLayoutMode.justified}
                    constrainMode={ConstrainMode.horizontalConstrained}
                    isHeaderVisible={true}
                />): (
                <ScrollablePane style = {{
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
                </ScrollablePane>)
            }

            <Panel
                headerText={selectedIncidentTitle}
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
                                    disabled={this._isFirst(selectedIncident)}
                                    onClick={() => this._selectPrevious(selectedIncident)} />

                                <IconButton
                                    iconProps={{ iconName: 'Down' }}
                                    title='Next'
                                    ariaLabel='Next item'
                                    disabled={this._isLast(selectedIncident)}
                                    onClick={() => this._selectNext(selectedIncident)} />
                            </>)}
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
                <IncidentDetails
                    id={selectedIncident}
                    onPublishingChange={this.handleItemChange}
                    onView={this.handleView}
                    onFavorite={this.handleFavorite}
                    onUpdateOrgTags={this.handleOrgTagsUpdate}
                    onUpdateParentTitle={this._updatePanelTitle}
                />
            </Panel>
        </>
        );
    }

    componentDidMount() {
        this._getIncidents();
    }

    private _getIncidents() {
        var incidents: IHealthEvent[] = [];
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
                    fetch('/api/issues?select=id,title,affectedServices,startDateTime,lastModifiedDateTime,isResolved,public,classificationDisplayName,statusDisplayName,originDisplayName,summary,serviceHealthHubState&expand=summary,serviceHealthHubViewpoint,organizationTags&filter=isResolved eq ' + (this.props.type === HealthReportType.active ? "false" : "true"),
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

                                for (const incident of sortedResult) {
                                    var start: Date = new Date(incident.startDateTime);
                                    var lastModified: Date = new Date(incident.lastModifiedDateTime);
                                    var summarySearchable: string | undefined = incident.summary ?
                                        incident.summary.contents?.map((s: any) => s.text)?.join(' ')
                                        : undefined;

                                    var orgTags: string[] = incident.organizationTags.map((ot: any) => ot.tag);

                                    incidents.push({
                                        id: incident.id,
                                        title: incident.title,
                                        service: incident.affectedServices,
                                        start: start,
                                        lastModified: lastModified,
                                        isResolved: incident.isResolved,
                                        published: incident.public,
                                        issueType: incident.classificationDisplayName,
                                        status: incident.statusDisplayName,
                                        origin: incident.originDisplayName,
                                        summary: incident.summary,
                                        summarySearchable: summarySearchable,
                                        favorite: incident.serviceHealthHubViewpoint?.favorited ? incident.serviceHealthHubViewpoint.favorited : false,
                                        serviceHealthHubState: incident.serviceHealthHubState == null ? "PENDING" : incident.serviceHealthHubState,
                                        serviceHealthHubViewpoint: incident.serviceHealthHubViewpoint,
                                        orgTags: orgTags
                                    });
                                }
                            }

                            try {
                                if (this.props.onDataLoaded)
                                    this.props.onDataLoaded(incidents)
                            }
                            catch(err)
                            {
                                this.setState({
                                    error: err.message
                                });
                            }

                            this._allItems = incidents;
                            var viewState = this._getViewTypeFilter(this._allItems);
                            var filteredItems: IHealthEvent[] = this.state.filter ? this.state.filter.filterItems(viewState) : viewState;

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

    componentDidUpdate(prevProps: any) {

        if (this.props.filterChangeToken && this.props.filterChangeToken !== this.state.filterChangeToken) {
            var viewState = this._getViewTypeFilter(this._allItems);
            var filteredItems: IHealthEvent[] = this.state.filter ? this.state.filter.filterItems(viewState) : viewState;

            if (this.props.onFilterChange)
                try
                {
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

    private _getViewTypeFilter(source: IHealthEvent[]): IHealthEvent[] {
        var events: IHealthEvent[] = [];
        Object.assign(events, source);

        if (this.props.collectionMode)
            switch (this.props.collectionMode) {
                case HealthItems.Inbox:
                    events = events.filter((e: IHealthEvent) => !e.serviceHealthHubViewpoint?.archived);
                    break;
                case HealthItems.Archived:
                    events = events.filter((e: IHealthEvent) => e.serviceHealthHubViewpoint?.archived);
                    break;
                case HealthItems.Favorites:
                    events = events.filter((e: IHealthEvent) => e.serviceHealthHubViewpoint?.favorite);
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

    private _getFilteredItems(items: IHealthEvent[]): IHealthEvent[] {
        var viewState = this._getViewTypeFilter(this._allItems);
        var filteredItems: IHealthEvent[] = this.state.filter ? this.state.filter.filterItems(viewState) : viewState;

        if (this.props.onFilterChange)
            this.props.onFilterChange(filteredItems.length);

        return filteredItems;
    }

    private _switchViewState(id: string): void {
        var item: IHealthEvent | undefined = this._allItems.find(m => m.id.toLowerCase() === id.toLowerCase());

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
        var item: IHealthEvent | undefined = this._allItems.find(m => m.id.toLowerCase() === id.toLowerCase());

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
        var item: IHealthEvent | undefined = this._allItems.find(m => m.id.toLowerCase() === id.toLowerCase());

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

    _updatePanelTitle(title: string) {
        this.setState({
            selectedIncidentTitle: title
        });
    }

    _onOpenDetailsPanel(id: string) {
        this.setState({
            selectedIncident: id,
            selectedIncidentTitle: this._allItems.find(m => m.id.toLowerCase() === id.toLowerCase())?.title!,
            disablePanelNavigation: false,
            isDetailsPanelOpen: true
        });
    }

    private _onDismisDetailsPanel = (): void => {
        this.setState({
            selectedIncident: "",
            isDetailsPanelOpen: false
        });
    }

    private _selectPrevious(id: string): void {
        if (this.state.items !== undefined) {
            const itemIndex: number = this.state.items.findIndex((i) => i.id === id);

            if (itemIndex > 0) {
                this.setState({
                    selectedIncident: this.state.items[itemIndex - 1].id,
                    selectedIncidentTitle: this._allItems.find(m => m.id.toLowerCase() === this.state.items[itemIndex - 1].id.toLowerCase())?.title!,
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
                    selectedIncident: this.state.items[itemIndex + 1].id,
                    selectedIncidentTitle: this._allItems.find(m => m.id.toLowerCase() === this.state.items[itemIndex + 1].id.toLowerCase())?.title!,
                    disablePanelNavigation: false
                });
            }
        }
    }

    private _isFirst(id: string): boolean {
        if (this.state.items !== undefined && this.state.items.length > 0) {
            const firstElement: IHealthEvent = this.state.items[0];
            return firstElement.id === id;
        } else
            return true;
    }

    private _isLast(id: string): boolean {
        if (this.state.items !== undefined && this.state.items.length > 0) {
            const lastElement: IHealthEvent = this.state.items[this.state.items.length - 1];
            return lastElement.id === id;
        } else
            return true;
    }
}