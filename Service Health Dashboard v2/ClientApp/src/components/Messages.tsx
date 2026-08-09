/* eslint-disable no-loop-func */
import * as React from 'react';
import { GlobalState } from './GlobalState';
import { NotificationState } from './Notifications';
import {
    IIconProps, ILinkStyleProps, ILinkStyles, Link, Text, SearchBox, FontIcon, IconButton, CommandBarButton,
    ResponsiveMode, Button, ITheme, DefaultPalette
} from '@fluentui/react';
import { Fabric, mergeStyles, mergeStyleSets } from '@fluentui/react';
import { DetailsList, DetailsListLayoutMode, Selection, SelectionMode, IColumn, IDetailsListStyles } from '@fluentui/react';
import { Panel, PanelType} from '@fluentui/react';
import { IDropdownOption, IContextualMenuItem } from '@fluentui/react';
import { CommandBar, ICommandBarItemProps, Pivot, PivotItem } from '@fluentui/react';
import { TooltipHost } from '@fluentui/react';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { IBreadcrumbItem } from '@fluentui/react';
import { M365Breadcrumb } from '@m365-admin/m365-breadcrumb';
import { DetailPageHeader } from '@m365-admin/detail-page';
import { ServiceComponent } from './ServiceNameComponent';
import { Tag } from './TagComponent';
import { MessageDetails } from './MessageDetails';
import { MessageBar, MessageBarType, ThemeProvider } from '@fluentui/react';
import { acquireAccessToken } from "../auth/AccessTokenHelper";
import { checkInTeams } from './auth/detectTeams';
import { AccessDenied } from "./AccessDenied";
import { InPageFilterSimple, FilterPill } from '@m365-admin/in-page-filter';

export interface IMessageListState {
    columns: IColumn[];
    items: IMessage[];
    archivedItems: IMessage[];
    servicesOptions: IDropdownOption[];
    tagOptions: IDropdownOption[];
    platformOptions: IDropdownOption[];
    orgTagOptions: IDropdownOption[];
    rolloutStatusOptions: IDropdownOption[];
    filterServicesOptions: IContextualMenuItem[];
    filterTagOptions: IContextualMenuItem[];
    filterRolloutStatusOptions: IContextualMenuItem[];
    filterLastUpdateOptions: IContextualMenuItem[];
    filterPublishedOptions: IContextualMenuItem[];
    filterPlatformOptions: IContextualMenuItem[];
    filterOrgTagOptions: IContextualMenuItem[];
    selection: Selection;
    selectionDetails: string;
    announcedMessage?: string;
    currentView: string | undefined;
    isDataLoaded: boolean;
    isPanelOpen: boolean;
    isFilterPanelOpen: boolean;
    selectedMessage: string;
    selectedMessageTitle: string;
    disablePanelNavigation: boolean;
    inTeams: boolean;
    error?: string;
    accessGranted?: boolean;
    theme?: any;
}

export interface IMessage {
    key: string;
    name: string;
    value: string;
    id: string;
    title: string;
    service: string[];
    lastUpdated: Date;
    actBy?: Date;
    majorUpdate: boolean;
    tags: string[];
    rolloutStatus?: string;
    platforms: string[];
    extendedProperties: any;
    serviceHealthHubViewpoint: any;
    orgTags: string[];
    published: boolean;
}

interface IMessageFilter {
    textFilter: string;
    serviceFilter: string[];
    tagFilter: string[];
    rolloutStatusFilter: string[];
    lastUpdateFilter: string[];
    publishedFilter: string[];
    platformFilter: string[];
    orgTagFilter: string[];
}

const searchIcon: IIconProps = { iconName: 'Search' };

export class MessagesList extends React.Component<{}, IMessageListState> {
    static contextType = GlobalState;
    private _selection: Selection;
    private _archiveSelection: Selection;
    private _allItems: IMessage[];
    private _breadcrumbItems: IBreadcrumbItem[] = [];
    private _filter: IMessageFilter = {
        textFilter: "",
        serviceFilter: [],
        tagFilter: [],
        rolloutStatusFilter: [],
        lastUpdateFilter: [],
        publishedFilter: [],
        platformFilter: [],
        orgTagFilter: []
    }

    constructor(props: {}) {
        super(props);

        this._allItems = [];

        this._selection = new Selection({
            selectionMode: SelectionMode.single
        });

        this._archiveSelection = new Selection({
            selectionMode: SelectionMode.single
        });

        const pipeFabricStyles = (p: ILinkStyleProps): ILinkStyles => ({
            root: {
                textDecoration: 'none',
                color: p.theme.semanticColors.bodyText
            },
        });

        const iconClass = mergeStyles({
            fontSize: 16,
            height: 16,
            width: 16,
            margin: "0 16px"
        });

        const iconClassSmall = mergeStyles({

        });

        const classNames = mergeStyleSets({
            deepSkyBlue: [{ color: 'deepskyblue' }, iconClass],
            greenYellow: [{ color: 'greenyellow' }, iconClass],
            salmon: [{ color: 'salmon' }, iconClass],
            black: [{ color: 'black' }, iconClass],
            incident: [{ color: 'red' }, iconClass],
            advisory: [{ color: DefaultPalette.blue }, iconClass],
            default: [{}, iconClass],
        });

        const classNamesSmall = mergeStyleSets({
            incident: [{ color: 'red' }, iconClassSmall],
            advisory: [{ color: DefaultPalette.blue }, iconClassSmall]
        });

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
                onRender: (item: IMessage) => {
                    return <div className="container" style={{ cursor: 'default' }} >
                        <div className="row" >
                            <div className="col" style={{ justifyContent: 'center', paddingLeft: '0px', paddingRight: '0px', marginRight: '3px' }} >
                                {item.published ? (<FontIcon iconName='Streaming' />) : ""}
                            </div>
                            <div className="col" style={{ justifyContent: 'center', paddingLeft: '0px', paddingRight: '3px' }} >
                                <FontIcon
                                    iconName={item.majorUpdate ? 'WarningSolid' : 'InfoSolid'}
                                    className={item.majorUpdate ? classNamesSmall.incident : classNamesSmall.advisory} />
                            </div>
                        </div>
                    </div>;
                },
                isPadded: true,
            },
            {
                key: 'clTitle',
                name: 'Title',
                fieldName: 'title',
                minWidth: 240,
                maxWidth: 500,
                isRowHeader: false,
                isResizable: true,
                onRender: (item: IMessage) => {
                    return <Link onClick={(event) => {
                        event.preventDefault();
                        this._onOpenPanel(item.id);
                    }} styles={pipeFabricStyles}>{item.serviceHealthHubViewpoint && item.serviceHealthHubViewpoint.viewed ?
                        item.title :
                        (<b>{item.title}</b>)}</Link >
                },
                data: 'string',
                isPadded: true,
            },
            {
                key: 'clService',
                name: 'Service',
                fieldName: 'service',
                minWidth: 100,
                maxWidth: 200,
                isResizable: true,
                isCollapsible: false,
                data: 'string',
                onColumnClick: this._onColumnClick,
                onRender: (item: IMessage) => {
                    return <TooltipHost content={item.service.join(", ")}><span>{
                        item.service.map((s): JSX.Element => (
                            <ServiceComponent name={s} />
                        ))
                    }</span></TooltipHost>;
                },
                isPadded: true,
            },
            {
                key: 'clLastUpdated',
                name: 'Last updated',
                fieldName: 'lastUpdated',
                minWidth: 60,
                maxWidth: 80,
                isResizable: true,
                isCollapsible: true,
                isSorted: true,
                isSortedDescending: true,
                sortAscendingAriaLabel: 'Sorted newer to older',
                sortDescendingAriaLabel: 'Sorted older to newer',
                data: 'string',
                onColumnClick: this._onColumnClick,
                isPadded: true,
                onRender: (item: IMessage) => {
                    return <span>{item.lastUpdated.toLocaleDateString()}</span>;
                },
            },
            {
                key: 'clRolloutStatus',
                name: 'Status for your org',
                fieldName: 'rolloutStatus',
                minWidth: 80,
                maxWidth: 120,
                isResizable: true,
                isCollapsible: true,
                isSorted: false,
                isSortedDescending: false,
                sortAscendingAriaLabel: 'Sorted newer to older',
                sortDescendingAriaLabel: 'Sorted older to newer',
                data: 'string',
                onColumnClick: this._onColumnClick,
                isPadded: true,
                onRender: (item: IMessage) => {
                    return <span>{item.rolloutStatus ? item.rolloutStatus : ""}</span>;
                },
            },
            {
                key: 'clMessageId',
                name: 'Message ID',
                fieldName: 'id',
                minWidth: 60,
                maxWidth: 80,
                isResizable: true,
                isCollapsible: true,
                data: 'string',
                onColumnClick: this._onColumnClick,
                isPadded: true,
            },
            {
                key: 'clTags',
                name: 'Tags',
                fieldName: 'tags',
                minWidth: 80,
                maxWidth: 300,
                isResizable: true,
                isCollapsible: true,
                data: 'string',
                isPadded: true,
                onRender: (item: IMessage) => {
                    return <TooltipHost content={item.tags.join(", ")}><span>{
                        item.tags.map((t) => (
                            <Tag name={t} key={t} />
                        ))
                    }</span></TooltipHost>;
                },
            },
            {
                key: 'clOrgTags',
                name: 'Org. tags',
                fieldName: 'orgTags',
                minWidth: 80,
                maxWidth: 300,
                isResizable: true,
                isCollapsible: true,
                data: 'string',
                isPadded: true,
                onRender: (item: IMessage) => {
                    return <TooltipHost content={item.orgTags.join(", ")}><span>{
                        item.orgTags.map((t) => (
                            <Tag name={t} key={t} />
                        ))
                    }</span></TooltipHost>;
                },
            },
            {
                key: 'clActBy',
                name: 'Act By',
                fieldName: 'actBy',
                minWidth: 60,
                maxWidth: 80,
                isResizable: true,
                isCollapsible: true,
                data: 'string',
                onColumnClick: this._onColumnClick,
                isPadded: true,
                onRender: (item: IMessage) => {
                    return <span>{item.actBy ? item.actBy.toLocaleDateString() : ""}</span>;
                },
            },
        ];

        const queryParams = new URLSearchParams(window.location.search);
        const id = queryParams.get('id');

        this.state = {
            items: this._allItems,
            archivedItems: [],
            columns: columns,
            servicesOptions: [],
            tagOptions: [],
            rolloutStatusOptions: [],
            platformOptions: [],
            orgTagOptions: [],
            filterServicesOptions: [],
            filterTagOptions: [],
            filterRolloutStatusOptions: [],
            filterPlatformOptions: [],
            filterOrgTagOptions: [],
            filterLastUpdateOptions: [
                { key: '7', text: '7 days', onClick: () => this._setLastUpdateFilterItem('7'), canCheck: true, checked: false },
                { key: '30', text: '30 days', onClick: () => this._setLastUpdateFilterItem('30'), canCheck: true, checked: false },
                { key: '90', text: '90 days', onClick: () => this._setLastUpdateFilterItem('90'), canCheck: true, checked: false },
                { key: '180', text: '180 days', onClick: () => this._setLastUpdateFilterItem('180'), canCheck: true, checked: false }
            ],
            filterPublishedOptions: [
                { key: 'published', text: 'Yes', onClick: () => this._setPublishedFilterItem('published'), canCheck: true, checked: false },
                { key: 'notpublished', text: 'No', onClick: () => this._setPublishedFilterItem('notpublished'), canCheck: true, checked: false }
            ],
            selection: this._selection,
            selectionDetails: this._getSelectionDetails(),
            announcedMessage: undefined,
            isDataLoaded: false,
            isPanelOpen: !(id === null || id === undefined),
            isFilterPanelOpen: false,
            selectedMessage: id === null || id === undefined ? "" : id!,
            selectedMessageTitle: "",
            disablePanelNavigation: !(id === null || id === undefined),
            currentView: "inbox",
            inTeams: checkInTeams(),
            error: undefined,
            accessGranted: undefined
        };

        this._updatePanelTitle = this._updatePanelTitle.bind(this);

        this._breadcrumbItems = this.state.inTeams ? [
            { text: 'Upcoming features', key: 'upcomigFeatures' },
            { text: 'Message center', key: 'messageCenter', isCurrentItem: true }
        ] :
            [
                { text: 'Home', key: 'home', isCurrentItem: false, href: '/' },
                { text: 'Upcoming features', key: 'upcomigFeatures' },
                { text: 'Message center', key: 'messageCenter', isCurrentItem: true }
            ];


    }

    handleItemChange = (itemId: string, published: boolean) => {
        var item = this._allItems.find(m => m.id.toLowerCase() === itemId.toLowerCase());
        if (item !== null && item !== undefined) {
            item.published = published;
        }
    }

    handleExtendedPropertiesChange = (itemId: string, extendedProperties: any) => {
        var item = this._allItems.find(m => m.id.toLowerCase() === itemId.toLowerCase());
        if (item !== null && item !== undefined) {
            item.extendedProperties = extendedProperties;
        }
        this._filterItems();
    }

    handleView = (itemId: string, viewed: boolean) => {
        var item = this._allItems.find(m => m.id.toLowerCase() === itemId.toLowerCase());
        if (item !== null && item !== undefined && item.serviceHealthHubViewpoint !== undefined) {
            item.serviceHealthHubViewpoint.viewed = viewed;
            this._filterItems();
        }
    }

    handleFavorite = (itemId: string, favorite: boolean) => {
        var item = this._allItems.find(m => m.id.toLowerCase() === itemId.toLowerCase());
        if (item !== null && item !== undefined && item.serviceHealthHubViewpoint !== undefined) {
            item.serviceHealthHubViewpoint.favorite = favorite;
            this._filterItems();
        }
    }

    handleArchive = (itemId: string, archived: boolean) => {
        var item = this._allItems.find(m => m.id.toLowerCase() === itemId.toLowerCase());
        if (item !== null && item !== undefined && item.serviceHealthHubViewpoint !== undefined) {
            item.serviceHealthHubViewpoint.archived = archived;
            this._filterItems();
        }
    }

    handleOrgTagsUpdate = (itemId: string, orgTags: string[]) => {
        var item = this._allItems.find(m => m.id.toLowerCase() === itemId.toLowerCase());
        if (item !== null && item !== undefined) {
            item.orgTags = orgTags;
            this._filterItems();
        }
    }

    public render() {
        const {
            columns, items, archivedItems, isDataLoaded, isPanelOpen, selectedMessage,
            filterServicesOptions, filterTagOptions, filterLastUpdateOptions, filterPublishedOptions,
            filterRolloutStatusOptions, currentView, filterPlatformOptions, filterOrgTagOptions,
            error, accessGranted, theme, selectedMessageTitle, disablePanelNavigation
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

        if (theme === undefined)
            return "";

        const gridStyles: Partial<IDetailsListStyles> = {
            root: {
                width: '101%',
                overflowX: 'hidden',
                overflowY: 'hidden',
                selectors: {
                    '& [role=grid]': {
                        display: 'flex',
                        flexDirection: 'column',
                        alignItems: 'start',
                        height: '70vh',
                    },
                },
            },
            headerWrapper: {
                flex: '0 0 auto',
            },
            contentWrapper: {
                flex: '1 1 auto',
                overflowY: 'auto',
                overflowX: 'hidden',
            },
        };

        const _commandBarItems: ICommandBarItemProps[] = [
        ];

        var itemCount: number = 0;

        switch (currentView) {
            case 'inbox':
                itemCount = items.length
                break;

            case 'archive':
                itemCount = archivedItems.length;
                break;

            default:
                itemCount = items.length;
        }

        const _commandBarFarItems: ICommandBarItemProps[] = [
            {
                key: 'itemsCount',
                onRender: () => {
                    return <div style={{ paddingTop: 4, paddingRight: 8 }}><Text variant='smallPlus'><b>{itemCount} item{itemCount !== 1 ? "s" : ""}</b></Text></div>;
                },
            },
            {
                key: 'searchBox',
                onRender: () => {
                    return <div style={{ paddingLeft: 8, paddingRight: 8 }}><SearchBox placeholder="Search" iconProps={searchIcon} onChange={this._onChangeSearchText} /></div>;
                },
            }
        ];

        return (
            <Fabric>

                <div className="container">
                    <div className="row">
                        <div className="col">
                            <M365Breadcrumb
                                items={this._breadcrumbItems}
                                style={{ marginBottom: '16px' }}
                                ariaLabel="Breadcrumb with items rendered as buttons"
                                overflowAriaLabel="More links"
                            />
                            <DetailPageHeader
                                title="Message Center"
                                description="Each message gives you a high-level overview of a planned change and how it may affect your users, and links out to more detailed information to help you prepare."
                            />
                        </div>
                    </div>

                    <div className="row">
                        <div className="col">
                            <Pivot
                                aria-label="Message"
                                linkFormat={'links'}
                                overflowBehavior={'menu'}
                                overflowAriaLabel="more items"
                                onLinkClick={this._onArchivePivotChange}
                            >

                                <PivotItem headerText="Inbox" id='inbox'>
                                    <div className="row">
                                        <div className="col">
                                            <div>
                                                <CommandBar items={_commandBarItems} farItems={_commandBarFarItems} />
                                            </div>
                                            <InPageFilterSimple
                                                label="Filters"
                                                addFilterText="Add filter"
                                                styles={{ label: { fontSize: '16px' } }}
                                                showResetFiltersButton={false}
                                            >
                                                <FilterPill
                                                    name="Services"
                                                    menuProps={{ items: filterServicesOptions, calloutProps: { calloutMaxHeight: 300 } }}
                                                    filterPillMenuProps={{ headerText: 'Services' }}
                                                    defaultValue={'All'}
                                                    value={this._filter.serviceFilter}
                                                    resetProps={{
                                                        onClick: () => { this._clearFilter() },
                                                        'aria-label': 'Clear services filters',
                                                    }}
                                                />
                                                <FilterPill
                                                    name="Platforms"
                                                    menuProps={{ items: filterPlatformOptions, calloutProps: { calloutMaxHeight: 300 } }}
                                                    filterPillMenuProps={{ headerText: 'Services' }}
                                                    defaultValue={'All'}
                                                    value={this._filter.platformFilter}
                                                    resetProps={{
                                                        onClick: () => { this._clearFilter() },
                                                        'aria-label': 'Clear platform filters',
                                                    }}
                                                />
                                                <FilterPill
                                                    name="Tags"
                                                    menuProps={{
                                                        items: filterTagOptions, calloutProps: { calloutMaxHeight: 300 }
                                                    }}
                                                    filterPillMenuProps={{ headerText: 'Tags' }}
                                                    defaultValue={'All'}
                                                    value={this._filter.tagFilter}
                                                    resetProps={{
                                                        onClick: () => { },
                                                        'aria-label': 'Clear tags filters',
                                                    }}
                                                />
                                                <FilterPill
                                                    name="Rollout status"
                                                    menuProps={{
                                                        items: filterRolloutStatusOptions, calloutProps: { calloutMaxHeight: 300 }
                                                    }}
                                                    filterPillMenuProps={{ headerText: 'Rollout status' }}
                                                    defaultValue={'All'}
                                                    value={this._filter.rolloutStatusFilter}
                                                    resetProps={{
                                                        onClick: () => { },
                                                        'aria-label': 'Clear rollout status filters',
                                                    }}
                                                />
                                                <FilterPill
                                                    name="Last update"
                                                    menuProps={{
                                                        items: filterLastUpdateOptions, calloutProps: { calloutMaxHeight: 300 }
                                                    }}
                                                    filterPillMenuProps={{ headerText: 'Last update' }}
                                                    defaultValue={'All'}
                                                    value={this._filter.lastUpdateFilter}
                                                    resetProps={{
                                                        onClick: () => { },
                                                        'aria-label': 'Clear last update filter',
                                                    }}
                                                />
                                                <FilterPill
                                                    name="Published"
                                                    menuProps={{
                                                        items: filterPublishedOptions, calloutProps: { calloutMaxHeight: 300 }
                                                    }}
                                                    filterPillMenuProps={{ headerText: 'Published' }}
                                                    defaultValue={'All'}
                                                    value={this._filter.publishedFilter}
                                                    resetProps={{
                                                        onClick: () => { },
                                                        'aria-label': 'Clear last update filter',
                                                    }}
                                                />
                                                {filterOrgTagOptions.length > 0 ? (
                                                    <FilterPill
                                                        name="Org. tag"
                                                        menuProps={{
                                                            items: filterOrgTagOptions, calloutProps: { calloutMaxHeight: 300 }
                                                        }}
                                                        filterPillMenuProps={{ headerText: 'Org. tag' }}
                                                        defaultValue={'All'}
                                                        value={this._filter.orgTagFilter}
                                                        resetProps={{
                                                            onClick: () => { },
                                                            'aria-label': 'Clear organization tag filter',
                                                        }}
                                                    />
                                                ) : (<></>)}
                                                <CommandBarButton
                                                    iconProps={{ iconName: 'Save' }}
                                                    text="Save filter"
                                                    onClick={() => { this._saveFilter() }}
                                                />
                                                <CommandBarButton
                                                    iconProps={{ iconName: 'ClearFilter' }}
                                                    text="Reset all"
                                                    onClick={() => { this._clearFilter() }}
                                                />
                                            </InPageFilterSimple>

                                            <DetailsList
                                                key='inbox'
                                                items={items}
                                                compact={true}
                                                columns={columns}
                                                selection={this._selection}
                                                selectionMode={SelectionMode.none}
                                                getKey={this._getKey}
                                                setKey="id"
                                                layoutMode={DetailsListLayoutMode.justified}
                                                isHeaderVisible={true}
                                                onItemInvoked={this._onItemInvoked}
                                                isSelectedOnFocus={true}
                                            />
                                        </div>
                                    </div>
                                </PivotItem>

                                <PivotItem headerText="Archive" id='archive'>
                                    <div className="row">
                                        <div className="col">
                                            <div>
                                                <CommandBar items={_commandBarItems} farItems={_commandBarFarItems} />
                                            </div>

                                            <InPageFilterSimple
                                                label="Filters"
                                                addFilterText="Add filter"
                                                styles={{ label: { fontSize: '16px' } }}
                                                showResetFiltersButton={false}
                                            >
                                                <FilterPill
                                                    name="Services"
                                                    menuProps={{ items: filterServicesOptions, calloutProps: { calloutMaxHeight: 300 } }}
                                                    filterPillMenuProps={{ headerText: 'Services' }}
                                                    defaultValue={'All'}
                                                    value={this._filter.serviceFilter}
                                                    resetProps={{
                                                        onClick: () => { this._clearFilter() },
                                                        'aria-label': 'Clear services filters',
                                                    }}
                                                />
                                                <FilterPill
                                                    name="Platforms"
                                                    menuProps={{ items: filterPlatformOptions, calloutProps: { calloutMaxHeight: 300 } }}
                                                    filterPillMenuProps={{ headerText: 'Services' }}
                                                    defaultValue={'All'}
                                                    value={this._filter.platformFilter}
                                                    resetProps={{
                                                        onClick: () => { this._clearFilter() },
                                                        'aria-label': 'Clear platform filters',
                                                    }}
                                                />
                                                <FilterPill
                                                    name="Tags"
                                                    menuProps={{
                                                        items: filterTagOptions, calloutProps: { calloutMaxHeight: 300 }
                                                    }}
                                                    filterPillMenuProps={{ headerText: 'Tags' }}
                                                    defaultValue={'All'}
                                                    value={this._filter.tagFilter}
                                                    resetProps={{
                                                        onClick: () => { },
                                                        'aria-label': 'Clear tags filters',
                                                    }}
                                                />
                                                <FilterPill
                                                    name="Rollout status"
                                                    menuProps={{
                                                        items: filterRolloutStatusOptions, calloutProps: { calloutMaxHeight: 300 }
                                                    }}
                                                    filterPillMenuProps={{ headerText: 'Rollout status' }}
                                                    defaultValue={'All'}
                                                    value={this._filter.rolloutStatusFilter}
                                                    resetProps={{
                                                        onClick: () => { },
                                                        'aria-label': 'Clear rollout status filters',
                                                    }}
                                                />
                                                <FilterPill
                                                    name="Last update"
                                                    menuProps={{
                                                        items: filterLastUpdateOptions, calloutProps: { calloutMaxHeight: 300 }
                                                    }}
                                                    filterPillMenuProps={{ headerText: 'Last update' }}
                                                    defaultValue={'All'}
                                                    value={this._filter.lastUpdateFilter}
                                                    resetProps={{
                                                        onClick: () => { },
                                                        'aria-label': 'Clear last update filter',
                                                    }}
                                                />
                                                <FilterPill
                                                    name="Published"
                                                    menuProps={{
                                                        items: filterPublishedOptions, calloutProps: { calloutMaxHeight: 300 }
                                                    }}
                                                    filterPillMenuProps={{ headerText: 'Published' }}
                                                    defaultValue={'All'}
                                                    value={this._filter.publishedFilter}
                                                    resetProps={{
                                                        onClick: () => { },
                                                        'aria-label': 'Clear last update filter',
                                                    }}
                                                />
                                                {filterOrgTagOptions.length > 0 ? (
                                                    <FilterPill
                                                        name="Org. tag"
                                                        menuProps={{
                                                            items: filterOrgTagOptions, calloutProps: { calloutMaxHeight: 300 }
                                                        }}
                                                        filterPillMenuProps={{ headerText: 'Org. tag' }}
                                                        defaultValue={'All'}
                                                        value={this._filter.orgTagFilter}
                                                        resetProps={{
                                                            onClick: () => { },
                                                            'aria-label': 'Clear organization tag filter',
                                                        }}
                                                    />
                                                ) : (<></>)}
                                                <CommandBarButton
                                                    iconProps={{ iconName: 'Save' }}
                                                    text="Save filter"
                                                    onClick={() => { this._saveFilter() }}
                                                />
                                                <CommandBarButton
                                                    iconProps={{ iconName: 'ClearFilter' }}
                                                    text="Reset all"
                                                    onClick={() => { this._clearFilter() }}
                                                />
                                            </InPageFilterSimple>

                                            <DetailsList
                                                key='archive'
                                                items={archivedItems}
                                                compact={true}
                                                columns={columns}
                                                selection={this._archiveSelection}
                                                selectionMode={SelectionMode.none}
                                                getKey={this._getKey}
                                                setKey="id"
                                                layoutMode={DetailsListLayoutMode.justified}
                                                isHeaderVisible={true}
                                                onItemInvoked={this._onItemInvoked}
                                                isSelectedOnFocus={true}
                                            />
                                        </div>
                                    </div>
                                </PivotItem>
                            </Pivot>
                        </div>
                    </div>
                </div>

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

                <Panel
                    headerText={selectedMessageTitle}
                    isOpen={isPanelOpen}
                    onDismiss={this._onDismisPanel}
                    type={PanelType.medium}
                    isLightDismiss
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
                            </>)}

                            <IconButton
                                iconProps={{ iconName: 'Cancel' }}
                                title='Close'
                                ariaLabel='Close panel'
                                onClick={() => this._onDismisPanel()} />

                            {defaultRender!(props)}
                        </div>
                    )
                    }
                >
                    <MessageDetails
                        id={selectedMessage}
                        onPublishingChange={this.handleItemChange}
                        onExtendedPropertiesChange={this.handleExtendedPropertiesChange}
                        onView={this.handleView}
                        onArchive={this.handleArchive}
                        onFavorite={this.handleFavorite}
                        onUpdateOrgTags={this.handleOrgTagsUpdate}
                        onUpdateParentTitle={this._updatePanelTitle}
                    />
                </Panel>
            </Fabric>
        );
    }

    public componentDidUpdate(previousProps: any, previousState: IMessageListState) {
        this._selection.setAllSelected(false);
    }

    componentDidMount() {
        var items: IMessage[] = [];
        var serviceList: IDropdownOption[] = [];
        var rolloutStatusList: IDropdownOption[] = [];
        var tagList: IDropdownOption[] = [];
        var platformList: IDropdownOption[] = [];
        var orgTagList: IDropdownOption[] = [];
        var filterServiceList: IContextualMenuItem[] = [];
        var filterTagList: IContextualMenuItem[] = [];
        var filterRolloutStatusList: IContextualMenuItem[] = [];
        var filterPlatformList: IContextualMenuItem[] = [];
        var filterOrgTagList: IContextualMenuItem[] = [];
        const requiredRoles: string[] = ['ServiceHealthReader', 'Communication.Write.All', 'Admin'];

        let globalState: any = this.context;
        var theme: ITheme = globalState.getTheme();

        this.setState({
            theme: theme
        });

        acquireAccessToken()
            .then((response) => {
                var tokenClaims: any = response.account?.idTokenClaims;
                const userRoles: any = tokenClaims?.roles;
                var userHasRequiredRole: boolean = userRoles.some((r: string) => requiredRoles.includes(r));

                this.setState({
                    accessGranted: userHasRequiredRole
                });

                if (userHasRequiredRole)
                    fetch('/api/Messages', { headers: response.idToken === "" ? {} : { 'Authorization': `Bearer ${response.idToken}` } })
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
                            for (const message of result) {
                                var featureStatusJson = message.details?.find((d: any) => d.name === 'FeatureStatusJson');
                                var rolloutStatus: string | undefined = undefined;

                                if (featureStatusJson !== undefined) {
                                    try {
                                        var featureStatus = JSON.parse(featureStatusJson.value);
                                        var fsKeys = Object.keys(featureStatus)
                                        var globalStateRoadmap = featureStatus[fsKeys[0]].find((gsr: any) => gsr.Platform === 'All');

                                        rolloutStatus = globalStateRoadmap && globalStateRoadmap.Status !== 'FeatureRolloutStatusNotSupported' ? globalStateRoadmap.Status : undefined
                                        if (rolloutStatus?.toLowerCase() === 'inrollout')
                                            rolloutStatus = 'Rolling out'

                                        if (rolloutStatus !== undefined) {
                                            var rs: string = rolloutStatus!;
                                            const rsKey: string = rs.replace(" ", "").toLowerCase();
                                            if (!rolloutStatusList.find(t => t.text.toLowerCase() === rs.toLowerCase())) {
                                                rolloutStatusList.push({ key: rsKey, text: rs });
                                                filterRolloutStatusList.push({
                                                    key: rsKey,
                                                    text: rs,
                                                    onClick: () => this._toggleRolloutStatusFilterItem(rsKey),
                                                    canCheck: true
                                                });
                                            }
                                        }
                                    } catch (e) {
                                        rolloutStatus = undefined
                                    }
                                }

                                var platformsListInt = message.details?.find((d: any) => d.name === 'Platforms');
                                var platforms = platformsListInt ? platformsListInt.value?.split(',') : undefined;
                                var plProcessed: string[] = [];
                                if (platforms) {
                                    for (const platform of platforms) {
                                        const p: string = platform.trim();
                                        plProcessed.push(p);
                                        if (!platformList.find(t => t.text.toLowerCase() === p.toLowerCase())) {
                                            platformList.push({ key: p.replace(" ", "").toLowerCase(), text: p });
                                            filterPlatformList.push({
                                                key: p.replace(" ", "").toLowerCase(),
                                                text: p,
                                                onClick: () => this._togglePlatformFilterItem(p.replace(" ", "").toLowerCase()),
                                                canCheck: true
                                            });
                                        }
                                    }
                                }

                                var orgTags: string[] = message.organizationTags.map((ot: any) => ot.tag);

                                items.push({
                                    key: message.id,
                                    name: message.title,
                                    value: message.id,
                                    id: message.id,
                                    title: message.title,
                                    majorUpdate: message.majorUpdate,
                                    service: message.services.length > 0 ? message.services : ["General announcement"],
                                    actBy: message.actionRequiredByDateTime ? new Date(message.actionRequiredByDateTime) : undefined,
                                    lastUpdated: new Date(message.lastModifiedDateTime),
                                    tags: message.tags ? message.tags.sort((a: string, b: string) => (a > b ? 1 : -1)) : [],
                                    published: message.public,
                                    rolloutStatus: rolloutStatus,
                                    platforms: plProcessed,
                                    extendedProperties: message.extendedProperties,
                                    serviceHealthHubViewpoint: message.serviceHealthHubViewpoint,
                                    orgTags: orgTags
                                });

                                for (const service of message.services) {
                                    if (!serviceList.find(s => s.text.toLowerCase() === service.toLowerCase())) {
                                        serviceList.push({ key: service.replace(" ", "").toLowerCase(), text: service });
                                        filterServiceList.push({
                                            key: service.replace(" ", "").toLowerCase(),
                                            text: service,
                                            onClick: () => this._toggleServiceFilterItem(service.replace(" ", "").toLowerCase()),
                                            canCheck: true
                                        });
                                    }
                                }

                                for (const tag of message.tags) {
                                    if (!tagList.find(t => t.text.toLowerCase() === tag.toLowerCase())) {
                                        tagList.push({ key: tag.replace(" ", "").toLowerCase(), text: tag });
                                        filterTagList.push({
                                            key: tag.replace(" ", "").toLowerCase(),
                                            text: tag,
                                            onClick: () => this._toggleTagFilterItem(tag.replace(" ", "").toLowerCase()),
                                            canCheck: true
                                        });
                                    }
                                }

                                for (const orgTag of orgTags) {
                                    if (!orgTagList.find(t => t.text.toLowerCase() === orgTag.toLowerCase())) {
                                        orgTagList.push({ key: orgTag.replace(" ", "").toLowerCase(), text: orgTag });
                                        filterOrgTagList.push({
                                            key: orgTag.replace(" ", "").toLowerCase(),
                                            text: orgTag,
                                            onClick: () => this._toggleOrgTagFilterItem(orgTag.replace(" ", "").toLowerCase()),
                                            canCheck: true
                                        });
                                    }
                                }

                                items = items.sort((a, b) => 0 - (a.lastUpdated > b.lastUpdated ? 1 : -1));
                                serviceList = serviceList.sort((a, b) => 0 - (a.text > b.text ? -1 : 1));
                                tagList = tagList.sort((a, b) => 0 - (a.text > b.text ? -1 : 1));
                                platformList = platformList.sort((a, b) => 0 - (a.text > b.text ? -1 : 1));
                                rolloutStatusList = rolloutStatusList.sort((a, b) => 0 - (a.text > b.text ? -1 : 1));
                                orgTagList = orgTagList.sort((a, b) => 0 - (a.text > b.text ? -1 : 1));
                                filterServiceList = filterServiceList.sort((a, b) => 0 - (a.text! > b.text! ? -1 : 1));
                                filterTagList = filterTagList.sort((a, b) => 0 - (a.text! > b.text! ? -1 : 1));
                                filterRolloutStatusList = filterRolloutStatusList.sort((a, b) => 0 - (a.text! > b.text! ? -1 : 1));
                                filterPlatformList = filterPlatformList.sort((a, b) => 0 - (a.text! > b.text! ? -1 : 1));
                                filterOrgTagList = filterOrgTagList.sort((a, b) => 0 - (a.text! > b.text! ? -1 : 1));
                            }
                        }).then(() => {
                            this._allItems = items;
                            this.setState({
                                servicesOptions: serviceList,
                                tagOptions: tagList,
                                filterServicesOptions: filterServiceList,
                                filterTagOptions: filterTagList,
                                filterRolloutStatusOptions: filterRolloutStatusList,
                                rolloutStatusOptions: rolloutStatusList,
                                platformOptions: platformList,
                                filterPlatformOptions: filterPlatformList,
                                orgTagOptions: orgTagList,
                                filterOrgTagOptions: filterOrgTagList,
                                isDataLoaded: true
                            });
                            this._loadFilter();
                            this._filterItems();
                        });
            }).catch((err) => {
                this.setState({
                    error: err.message
                });
            });
    }

    private _getKey(item: any, index?: number): string {
        return item.key;
    }

    private _getItemsCollection(): IMessage[] {
        let collection: IMessage[];

        switch (this.state.currentView) {
            case 'inbox':
                collection = this.state.items;
                break;
            case 'archive':
                collection = this.state.archivedItems;
                break;
            default:
                collection = this.state.items;
        }

        return collection;
    }

    private _getListSelectionObject(): Selection {
        let selection: Selection;

        switch (this.state.currentView) {
            case 'inbox':
                selection = this._selection;
                break;
            case 'archive':
                selection = this._archiveSelection;
                break;
            default:
                selection = this._selection;
        }

        return selection;
    }

    private _selectPrevious(id: string): void {
        const collection: IMessage[] = this._getItemsCollection();
        const selection: Selection = this._getListSelectionObject();

        const itemIndex: number = collection.findIndex((i) => i.id === id);

        if (itemIndex > 0) {
            selection.setIndexSelected(itemIndex, false, false);
            selection.setIndexSelected(itemIndex - 1, true, false);
            this.setState({
                selectedMessage: collection[itemIndex - 1].id,
                selectedMessageTitle: this._allItems.find(m => m.id.toLowerCase() === collection[itemIndex - 1].id.toLowerCase())?.title!,
                disablePanelNavigation: false
            });
        }
    }

    private _selectNext(id: string): void {
        const collection: IMessage[] = this._getItemsCollection();
        const selection: Selection = this._getListSelectionObject();

        const itemIndex: number = collection.findIndex((i) => i.id === id);

        if (itemIndex < collection.length) {
            selection.setIndexSelected(itemIndex, false, false);
            selection.setIndexSelected(itemIndex + 1, true, false);
            this.setState({
                selectedMessage: collection[itemIndex + 1].id,
                selectedMessageTitle: this._allItems.find(m => m.id.toLowerCase() === collection[itemIndex + 1].id.toLowerCase())?.title!,
                disablePanelNavigation: false
            });
        }
    }

    private _isFirst(id: string): boolean {
        const collection: IMessage[] = this._getItemsCollection();

        if (collection.length === 0)
            return true;

        const firstElement: IMessage = collection[0];
        return firstElement.id === id;
    }

    private _isLast(id: string): boolean {
        const collection: IMessage[] = this._getItemsCollection();

        if (collection.length === 0)
            return true;

        const lastElement: IMessage = collection[collection.length - 1];
        return lastElement.id === id;
    }

    private _filterItemsInt(archived: boolean): IMessage[] {
        var filteredItems: IMessage[];
        var textFilter: string = this._filter.textFilter;
        var servicesFilter: string[] = this._filter.serviceFilter;
        var tagFilter: string[] = this._filter.tagFilter;
        var rolloutStatusFilter: string[] = this._filter.rolloutStatusFilter;
        var platformFilter: string[] = this._filter.platformFilter;
        var orgTagFilter: string[] = this._filter.orgTagFilter;
        var lastUpdateFilter: string | undefined = this._filter.lastUpdateFilter.length > 0 ?
            this.state.filterLastUpdateOptions.find(o => o.text === this._filter.lastUpdateFilter[0])?.key : "";
        var publishedFilter: string | undefined = this._filter.publishedFilter.length > 0 ?
            this.state.filterPublishedOptions.find(o => o.text === this._filter.publishedFilter[0])?.key : "";

        filteredItems = archived ?
            this._allItems.filter(i =>
                i.serviceHealthHubViewpoint !== undefined &&
                i.serviceHealthHubViewpoint !== null &&
                i.serviceHealthHubViewpoint.archived).slice() :
            this._allItems.filter(i =>
                !(i.serviceHealthHubViewpoint !== undefined &&
                    i.serviceHealthHubViewpoint !== null &&
                    i.serviceHealthHubViewpoint.archived)).slice();

        filteredItems = textFilter !== "" ? filteredItems.filter(
            i => i.title.toLowerCase().indexOf(textFilter.toLowerCase()) > -1 ||
                i.id.toLowerCase().indexOf(textFilter.toLowerCase()) > -1 ||
                i.service.find(s => s.toLowerCase().indexOf(textFilter.toLowerCase()) > -1) !== undefined ||
                i.tags.find(t => t.toLowerCase().indexOf(textFilter.toLowerCase()) > -1) !== undefined ||
                i.rolloutStatus === textFilter ||
                i.platforms.find(t => t.toLowerCase().indexOf(textFilter.toLowerCase()) > -1) !== undefined
        ) : filteredItems;

        if (servicesFilter.length > 0) {
            filteredItems = filteredItems.filter(i => servicesFilter.some(s => i.service.indexOf(s) >= 0));
        }

        if (tagFilter.length > 0) {
            filteredItems = filteredItems.filter(i => tagFilter.some(t => i.tags.indexOf(t) >= 0));
        }

        if (rolloutStatusFilter.length > 0) {
            filteredItems = filteredItems.filter(i => rolloutStatusFilter.some(t => i.rolloutStatus === t));
        }

        if (platformFilter.length > 0) {
            filteredItems = filteredItems.filter(i => platformFilter.some(t => i.platforms.indexOf(t) >= 0));
        }

        if (orgTagFilter.length > 0) {
            filteredItems = filteredItems.filter(i => orgTagFilter.some(t => i.orgTags.indexOf(t) >= 0));
        }

        if (publishedFilter && publishedFilter !== "") {
            switch (publishedFilter) {
                case "published":
                    filteredItems = filteredItems.filter(i => i.published === true);
                    break;

                case "notpublished":
                    filteredItems = filteredItems.filter(i => i.published === undefined || i.published === false);
                    break;
            }
        }

        if (lastUpdateFilter && lastUpdateFilter != "") {
            var pastDaysFilter: number = parseInt(lastUpdateFilter);
            var date = new Date();
            date = new Date(date.setDate(date.getDate() - pastDaysFilter));
            filteredItems = filteredItems.filter(i => i.lastUpdated >= date);
        }

        return filteredItems
    }

    private _filterItems(): void {
        var filteredInbox: IMessage[] = this._filterItemsInt(false);
        var filteredArchive: IMessage[] = this._filterItemsInt(true);

        this.setState({
            items: filteredInbox,
            archivedItems: filteredArchive
        });
    }

    private _onChangeSearchText = (ev?: React.ChangeEvent<HTMLInputElement>, newValue?: string): void => {
        this._filter.textFilter = newValue ? newValue : "";
        this._filterItems();
    };

    private _toggleServiceFilterItem(key: string): void {
        var svcFilter: string[] = this._filter.serviceFilter;
        var changed: boolean = false;
        var option: IContextualMenuItem | undefined = this.state.filterServicesOptions.find((o) => o.key === key);

        if (option !== undefined) {
            option.checked = !option.checked;

            if (option.checked) {
                if (svcFilter.filter(s => s.toLowerCase() === option!.text!.toLowerCase()).length === 0) {
                    svcFilter.push(option.text!);
                    changed = true;
                }
            } else {
                var idx: number = svcFilter.indexOf(option!.text!);
                if (idx > -1) {
                    svcFilter.splice(idx, 1);
                    changed = true;
                }
            }
        }

        if (changed) {
            this._filter.serviceFilter = svcFilter;
            this._filterItems();
        }
    }

    private _toggleTagFilterItem = (key: string): void => {
        var tFilter: string[] = this._filter.tagFilter;
        var changed: boolean = false;
        var option: IContextualMenuItem | undefined = this.state.filterTagOptions.find((o) => o.key === key);

        if (option !== undefined) {
            option.checked = !option.checked;

            if (option.checked) {
                if (tFilter.filter(s => s.toLowerCase() === option!.text!.toLowerCase()).length === 0) {
                    tFilter.push(option!.text!);
                    changed = true;
                }
            } else {
                var idx: number = tFilter.indexOf(option!.text!);
                if (idx > -1) {
                    tFilter.splice(idx, 1);
                    changed = true;
                }
            }
        }

        if (changed) {
            this._filter.tagFilter = tFilter;
            this._filterItems();
        }
    }

    private _toggleRolloutStatusFilterItem = (key: string): void => {
        var tFilter: string[] = this._filter.rolloutStatusFilter;
        var changed: boolean = false;
        var option: IContextualMenuItem | undefined = this.state.filterRolloutStatusOptions.find((o) => o.key === key);

        if (option !== undefined) {
            option.checked = !option.checked;

            if (option.checked) {
                if (tFilter.filter(s => s.toLowerCase() === option!.text!.toLowerCase()).length === 0) {
                    tFilter.push(option!.text!);
                    changed = true;
                }
            } else {
                var idx: number = tFilter.indexOf(option!.text!);
                if (idx > -1) {
                    tFilter.splice(idx, 1);
                    changed = true;
                }
            }
        }

        if (changed) {
            this._filter.rolloutStatusFilter = tFilter;
            this._filterItems();
        }
    }

    private _togglePlatformFilterItem = (key: string): void => {
        var tFilter: string[] = this._filter.platformFilter;
        var changed: boolean = false;
        var option: IContextualMenuItem | undefined = this.state.filterPlatformOptions.find((o) => o.key === key);

        if (option !== undefined) {
            option.checked = !option.checked;

            if (option.checked) {
                if (tFilter.filter(s => s.toLowerCase() === option!.text!.toLowerCase()).length === 0) {
                    tFilter.push(option!.text!);
                    changed = true;
                }
            } else {
                var idx: number = tFilter.indexOf(option!.text!);
                if (idx > -1) {
                    tFilter.splice(idx, 1);
                    changed = true;
                }
            }
        }

        if (changed) {
            this._filter.platformFilter = tFilter;
            this._filterItems();
        }
    }

    private _toggleOrgTagFilterItem = (key: string): void => {
        var tFilter: string[] = this._filter.orgTagFilter;
        var changed: boolean = false;
        var option: IContextualMenuItem | undefined = this.state.filterOrgTagOptions.find((o) => o.key === key);

        if (option !== undefined) {
            option.checked = !option.checked;

            if (option.checked) {
                if (tFilter.filter(s => s.toLowerCase() === option!.text!.toLowerCase()).length === 0) {
                    tFilter.push(option!.text!);
                    changed = true;
                }
            } else {
                var idx: number = tFilter.indexOf(option!.text!);
                if (idx > -1) {
                    tFilter.splice(idx, 1);
                    changed = true;
                }
            }
        }

        if (changed) {
            this._filter.orgTagFilter = tFilter;
            this._filterItems();
        }
    }

    private _setLastUpdateFilterItem(key: string): void {
        var lastUpdateFilter: string[] = this._filter.lastUpdateFilter;
        var changed: boolean = false;
        var option: IContextualMenuItem | undefined = this.state.filterLastUpdateOptions.find((o) => o.key === key);

        if (option !== undefined) {
            option.checked = !option.checked;

            if (option.checked) {
                lastUpdateFilter = [];
                lastUpdateFilter.push(option.text!);
                changed = true;
            } else {
                lastUpdateFilter = [];
                changed = true;
            }
        }

        for (const filterOption of this.state.filterLastUpdateOptions)
            if (filterOption.key !== key)
                filterOption.checked = false;

        if (changed) {
            this._filter.lastUpdateFilter = lastUpdateFilter;
            this._filterItems();
        }
    }

    private _setPublishedFilterItem(key: string): void {
        var publishedFilter: string[] = this._filter.publishedFilter;
        var changed: boolean = false;
        var option: IContextualMenuItem | undefined = this.state.filterPublishedOptions.find((o) => o.key === key);

        if (option !== undefined) {
            option.checked = !option.checked;

            if (option.checked) {
                publishedFilter = [];
                publishedFilter.push(option.text!);
                changed = true;
            } else {
                publishedFilter = [];
                changed = true;
            }
        }

        for (const filterOption of this.state.filterPublishedOptions)
            if (filterOption.key !== key)
                filterOption.checked = false;

        if (changed) {
            this._filter.publishedFilter = publishedFilter;
            this._filterItems();
        }
    }

    private _resetFilterOptions(options: IContextualMenuItem[]): void {
        for (const option of options)
            option.checked = false;
    }

    private _clearFilter(): void {
        this._filter.serviceFilter = [];
        this._filter.tagFilter = [];
        this._filter.lastUpdateFilter = [];
        this._filter.publishedFilter = [];
        this._filter.rolloutStatusFilter = [];
        this._filter.platformFilter = [];
        this._filter.orgTagFilter = [];

        this._resetFilterOptions(this.state.filterServicesOptions);
        this._resetFilterOptions(this.state.filterTagOptions);
        this._resetFilterOptions(this.state.filterLastUpdateOptions);
        this._resetFilterOptions(this.state.filterPublishedOptions);
        this._resetFilterOptions(this.state.filterRolloutStatusOptions);
        this._resetFilterOptions(this.state.filterPlatformOptions);
        this._resetFilterOptions(this.state.filterOrgTagOptions);

        this._filterItems();
    }

    private _setFilterOptions(filter: any, options: IContextualMenuItem[]) {
        for (const option of options)
            option.checked = filter.find((o: any) => o === option.text) !== undefined;
    }

    private _loadFilter(): void {
        let globalState: any = this.context;
        var messageCenterFilter = globalState.getProfileProperty("messageCenterFilter");
        if (messageCenterFilter && messageCenterFilter != null) {
            if (messageCenterFilter.lastUpdateFilter)
                this._filter.lastUpdateFilter = messageCenterFilter.lastUpdateFilter;

            if (messageCenterFilter.publishedFilter)
                this._filter.publishedFilter = messageCenterFilter.publishedFilter;

            if (messageCenterFilter.serviceFilter)
                this._filter.serviceFilter = messageCenterFilter.serviceFilter.filter(
                    (f: string) => this.state.filterServicesOptions.find((c: IContextualMenuItem) => c.text?.toLowerCase().trim() === f?.toLowerCase().trim()));

            if (messageCenterFilter.tagFilter)
                this._filter.tagFilter = messageCenterFilter.tagFilter.filter(
                    (f: string) => this.state.filterTagOptions.find((c: IContextualMenuItem) => c.text?.toLowerCase().trim() === f?.toLowerCase().trim()));

            if (messageCenterFilter.rolloutStatusFilter)
                this._filter.rolloutStatusFilter = messageCenterFilter.rolloutStatusFilter.filter(
                    (f: string) => this.state.filterRolloutStatusOptions.find((c: IContextualMenuItem) => c.text?.toLowerCase().trim() === f?.toLowerCase().trim()));

            if (messageCenterFilter.platformFilter)
                this._filter.platformFilter = messageCenterFilter.platformFilter.filter(
                    (f: string) => this.state.filterPlatformOptions.find((c: IContextualMenuItem) => c.text?.toLowerCase().trim() === f?.toLowerCase().trim()));

            if (messageCenterFilter.orgTagFilter)
                this._filter.orgTagFilter = messageCenterFilter.orgTagFilter.filter(
                    (ot: string) => this.state.filterOrgTagOptions.find((c: IContextualMenuItem) => c.text?.toLowerCase().trim() === ot?.toLowerCase().trim()));

            this._filter.textFilter = '';
        }

        this._setFilterOptions(this._filter.lastUpdateFilter, this.state.filterLastUpdateOptions);
        this._setFilterOptions(this._filter.publishedFilter, this.state.filterPublishedOptions);
        this._setFilterOptions(this._filter.serviceFilter, this.state.filterServicesOptions);
        this._setFilterOptions(this._filter.tagFilter, this.state.filterTagOptions);
        this._setFilterOptions(this._filter.rolloutStatusFilter, this.state.filterRolloutStatusOptions);
        this._setFilterOptions(this._filter.platformFilter, this.state.filterPlatformOptions);
        this._setFilterOptions(this._filter.orgTagFilter, this.state.filterOrgTagOptions);
    }

    private _saveFilter(): void {
        let globalState: any = this.context;
        globalState.setProfileProperty("messageCenterFilter", this._filter);
        globalState.setToastNotification({
            id: '0',
            title: 'Saved',
            message: 'Message Center filter preferences saved.',
            time: new Date(),
            state: NotificationState.Information,
            active: true
        });
    }

    private _onItemInvoked(item: any): void {

    }

    private _getSelectionDetails(): string {
        const selectionCount = this._selection.getSelectedCount();

        switch (selectionCount) {
            case 0:
                return 'No items selected';
            case 1:
                return '1 item selected';
            default:
                return `${selectionCount} items selected`;
        }
    }

    private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
        const { columns, items } = this.state;
        const newColumns: IColumn[] = columns.slice();
        const currColumn: IColumn = newColumns.filter(currCol => column.key === currCol.key)[0];
        newColumns.forEach((newCol: IColumn) => {
            if (newCol === currColumn) {
                currColumn.isSortedDescending = !currColumn.isSortedDescending;
                currColumn.isSorted = true;
                this.setState({
                    announcedMessage: `${currColumn.name} is sorted ${currColumn.isSortedDescending ? 'descending' : 'ascending'
                        }`,
                });
            } else {
                newCol.isSorted = false;
                newCol.isSortedDescending = true;
            }
        });
        const newItems = _copyAndSort(items, currColumn.fieldName!, currColumn.isSortedDescending);
        this.setState({
            columns: newColumns,
            items: newItems,
        });
    };

    _onOpenPanel(id: string) {
        this.setState({
            selectedMessage: id,
            selectedMessageTitle: this._allItems.find(m => m.id.toLowerCase() === id.toLowerCase())?.title!,
            disablePanelNavigation: false,
            isPanelOpen: true
        });
    }

    _updatePanelTitle(title: string) {
        this.setState({
            selectedMessageTitle: title
        });
    }

    private _onDismisPanel = (): void => {
        this.setState({
            selectedMessage: "",
            selectedMessageTitle: "",
            isPanelOpen: false
        });
    }

    _onOpenFilterPanel() {
        this.setState({
            isFilterPanelOpen: true
        });
    }

    _onArchivePivotChange = (item?: PivotItem, ev?: React.MouseEvent<HTMLElement>): void => {
        if (item)
            this.setState({
                currentView: item.props.id
            });
    }
}

function _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
    const key = columnKey as keyof T;
    return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
}