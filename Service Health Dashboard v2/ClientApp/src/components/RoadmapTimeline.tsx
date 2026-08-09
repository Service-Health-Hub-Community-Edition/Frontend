import * as React from 'react';
import { GlobalState } from './GlobalState';
import { NotificationState } from './Notifications';
import { IIconProps, ILinkStyleProps, ILinkStyles, Link, Text, SearchBox, ISliderStyles } from '@fluentui/react';
import { Fabric } from '@fluentui/react';
import { Panel, PanelType } from '@fluentui/react';
import { Selection } from '@fluentui/react';
import { DetailsList, IColumn, SelectionMode, DetailsListLayoutMode, ITheme } from '@fluentui/react';
import { CommandBar, ICommandBarItemProps, IconButton, IContextualMenuItem } from '@fluentui/react';
import { TooltipHost, Pivot, PivotItem } from '@fluentui/react';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { CommandBarButton } from '@fluentui/react';
import { ITag } from '@fluentui/react';
import { IBreadcrumbItem } from '@fluentui/react';
import { M365Breadcrumb } from '@m365-admin/m365-breadcrumb';
import { DetailPageHeader } from '@m365-admin/detail-page';
import { InPageFilterSimple, FilterPill } from '@m365-admin/in-page-filter';
import { DefaultPalette, ThemeProvider } from '@fluentui/react';
import { MessageBar, MessageBarType } from '@fluentui/react';
import { Tag, TagViewState } from './TagComponent';
import { ServiceComponent } from './ServiceNameComponent';
import { acquireAccessToken } from "../auth/AccessTokenHelper";
import { checkInTeams } from './auth/detectTeams';
import { AccessDenied } from "./AccessDenied";
import { IRoadmapItem, IRoadmapTag, IMessageLink, RoadmapDetails } from './RoadmapDetails';

export interface IRoadmapTimelineState {
    items: IRoadmapItem[];
    archivedItems: IRoadmapItem[];
    tags: ITag[];
    filterProductsOptions: IContextualMenuItem[];
    filterCloudInstancesOptions: IContextualMenuItem[];
    filterPlatformsOptions: IContextualMenuItem[];
    filterReleasePhaseOptions: IContextualMenuItem[];
    selectedItem?: IRoadmapItem;
    selectedItemTitle: string;
    disablePanelNavigation: boolean;
    currentView: string | undefined;
    isDataLoaded: boolean;
    isDetailsPanelOpen: boolean;
    selectionDetails: string;
    inTeams: boolean;
    columns: IColumn[];
    announcedMessage: string;
    error?: string;
    accessGranted?: boolean;
    theme?: any;
}

interface IRoadmapFilter {
    textFilter: string;
    productsFilter: string[];
    cloudInstancesFilter: string[];
    platformsFilter: string[];
    releasePhaseFilter: string[];
}

const searchIcon: IIconProps = { iconName: 'Search' };

export class Roadmap extends React.Component<{}, IRoadmapTimelineState> {
    static contextType = GlobalState;
    private _selection: Selection;
    private _allItems: IRoadmapItem[] = [];
    private _breadcrumbItems: IBreadcrumbItem[] = [];
    private _filter: IRoadmapFilter = {
        textFilter: "",
        productsFilter: [],
        cloudInstancesFilter: [],
        platformsFilter: [],
        releasePhaseFilter: []
    };

    customAction: any = React.createRef();

    constructor(props: {}) {
        super(props);

        const pipeFabricStyles = (p: ILinkStyleProps): ILinkStyles => ({
            root: {
                textDecoration: 'none',
                color: p.theme.semanticColors.bodyText,
                fontWeight: '600'
            },
        });

        const columns: IColumn[] = [
            {
                key: 'clTitle',
                name: 'Title',
                fieldName: 'title',
                minWidth: 240,
                maxWidth: 500,
                isRowHeader: true,
                isResizable: true,
                onRender: (item: IRoadmapItem) => {
                    return <Link onClick={(event) => {
                        event.preventDefault();
                        this._onOpenDetailsPanel(item);
                    }} styles={pipeFabricStyles}>{item.title}</Link >;
                },
                data: 'string',
                isPadded: true,
            },
            {
                key: 'clAvailableFrom',
                name: 'Availability',
                fieldName: 'availabilityDate',
                minWidth: 60,
                maxWidth: 100,
                isResizable: true,
                isCollapsible: false,
                data: 'string',
                isPadded: true,
            },
            {
                key: 'clPlatforms',
                name: 'Platforms',
                fieldName: 'platforms',
                minWidth: 100,
                maxWidth: 200,
                isResizable: true,
                isCollapsible: false,
                data: 'string',
                onRender: (item: IRoadmapItem) => {
                    return <TooltipHost content={item.products.join(", ")}><span>{
                        item.products.map((s) => (
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
                onRender: (item: IRoadmapItem) => {
                    return <span>{item.lastUpdated.toLocaleDateString()}</span>;
                },
            },
            {
                key: 'clMessageId',
                name: 'Id',
                fieldName: 'id',
                minWidth: 60,
                maxWidth: 80,
                isResizable: true,
                isCollapsible: true,
                data: 'string',
                isPadded: true,
            },
            {
                key: 'clReleasePhase',
                name: 'Phase',
                fieldName: 'releasePhase',
                minWidth: 60,
                maxWidth: 300,
                isResizable: true,
                isCollapsible: true,
                data: 'string',
                isPadded: true,
                onRender: (item: IRoadmapItem) => {
                    return <TooltipHost content={item.releasePhase.join(", ")}><span>{
                        item.releasePhase.map((t) => (
                            <Tag name={t} viewState={TagViewState.TextOnly} key={t} />
                        ))
                    }</span></TooltipHost>;
                },
            }
        ];

        const queryParams = new URLSearchParams(window.location.search);
        const id = queryParams.get('id');

        this.state = {
            items: this._allItems,
            archivedItems: [],
            isDataLoaded: false,
            isDetailsPanelOpen: !(id === null || id === undefined),
            selectionDetails: "",
            currentView: "inbox",
            selectedItem: id === null || id === undefined ? undefined : {
                id: id,
                title: "",
                status: "",
                links: [],
                availabilityFrom: new Date(),
                availabilityTo: new Date(),
                availabilityDate: "",
                publicPreviewFrom: new Date(),
                publicPreviewTo: new Date(),
                publicPreviewDate: "",
                description: "",
                categories: [],
                tags: [],
                cloudInstances: [],
                products: [],
                releasePhase: [],
                platforms: [],
                published: new Date(),
                lastUpdated: new Date(),
                workItem: "",
                extendedProperties: null
            },
            selectedItemTitle: "",
            disablePanelNavigation: !(id === null || id === undefined),
            tags: [],
            filterProductsOptions: [],
            filterCloudInstancesOptions: [],
            filterPlatformsOptions: [],
            filterReleasePhaseOptions: [],
            inTeams: checkInTeams(),
            columns: columns,
            announcedMessage: '',
            error: undefined,
            accessGranted: undefined
        };

        this._updatePanelTitle = this._updatePanelTitle.bind(this);

        this._breadcrumbItems = this.state.inTeams ? [
            { text: 'Upcoming features', key: 'upcomigFeatures' },
            { text: 'Roadmap', key: 'roadmap', isCurrentItem: true }
        ] : 
            [
                { text: 'Home', key: 'home', isCurrentItem: false, href: '/' },
                { text: 'Upcoming features', key: 'upcomigFeatures' },
                { text: 'Roadmap', key: 'roadmap', isCurrentItem: true }
            ];

        this._selection = new Selection({
            onSelectionChanged: () => {
                this.setState({
                    selectionDetails: this._getSelectionDetails(),
                });
            },
        });
    }

    handleExtendedPropertiesChange = (itemId: string, extendedProperties: any) => {
        var item = this._allItems.find(m => m.id.toLowerCase() === itemId.toLowerCase());
        if (item !== null && item !== undefined) {
            item.extendedProperties = extendedProperties;
        }
        this._filterItems();
    }

    public render() {
        const {
            items, archivedItems, isDataLoaded, isDetailsPanelOpen,
            filterProductsOptions, filterCloudInstancesOptions, filterPlatformsOptions, filterReleasePhaseOptions,
            currentView, selectedItem, selectedItemTitle, disablePanelNavigation, columns, error, accessGranted, theme
        } = this.state;

        if (accessGranted === undefined) {
            return (<div />);
        } else {
            if (accessGranted === false) {
                return (<div>
                            <AccessDenied />
                    </div>);
            }
        }

        if (theme === undefined)
            return "";

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

        const _commandBarItems: ICommandBarItemProps[] = [
        ];

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
                                title="Microsoft 365 Roadmap"
                                description="The Microsoft 365 Roadmap lists updates that are currently planned for applicable subscribers. Please note that the information in this list is on as-is basis and may change at any time."
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
                                                name="Product"
                                                menuProps={{ items: filterProductsOptions, calloutProps: { calloutMaxHeight: 300 } }}
                                                filterPillMenuProps={{ headerText: 'Product' }}
                                                defaultValue={'All'}
                                                value={this._filter.productsFilter}
                                                resetProps={{
                                                    onClick: () => { },
                                                    'aria-label': 'Clear product filters',
                                                }}

                                            />
                                            <FilterPill
                                                name="Cloud instance"
                                                menuProps={{
                                                    items: filterCloudInstancesOptions, calloutProps: { calloutMaxHeight: 300 }
                                                }}
                                                filterPillMenuProps={{ headerText: 'Cloud instance' }}
                                                defaultValue={'All'}
                                                value={this._filter.cloudInstancesFilter}
                                                resetProps={{
                                                    onClick: () => { },
                                                    'aria-label': 'Clear cloud instance filters',
                                                }}
                                            />
                                            <FilterPill
                                                name="Platform"
                                                menuProps={{
                                                    items: filterPlatformsOptions, calloutProps: { calloutMaxHeight: 300 }
                                                }}
                                                filterPillMenuProps={{ headerText: 'Platform' }}
                                                defaultValue={'All'}
                                                value={this._filter.platformsFilter}
                                                resetProps={{
                                                    onClick: () => { },
                                                    'aria-label': 'Clear platform filter',
                                                }}
                                            />
                                            <FilterPill
                                                name="Release phase"
                                                menuProps={{
                                                    items: filterReleasePhaseOptions, calloutProps: { calloutMaxHeight: 300 }
                                                }}
                                                filterPillMenuProps={{ headerText: 'Release phase' }}
                                                defaultValue={'All'}
                                                value={this._filter.releasePhaseFilter}
                                                resetProps={{
                                                    onClick: () => { },
                                                    'aria-label': 'Clear release phase filter',
                                                }}
                                            />
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
                                        items={items}
                                        compact={false}
                                        columns={columns}
                                        selectionMode={SelectionMode.none}
                                        getKey={this._getKey}
                                        setKey="id"
                                        layoutMode={DetailsListLayoutMode.justified}
                                        isHeaderVisible={true}
                                    />
                                </PivotItem>

                                <PivotItem headerText="Archive" id='archive'>
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
                                                name="Product"
                                                menuProps={{ items: filterProductsOptions, calloutProps: { calloutMaxHeight: 300 } }}
                                                filterPillMenuProps={{ headerText: 'Product' }}
                                                defaultValue={'All'}
                                                value={this._filter.productsFilter}
                                                resetProps={{
                                                    onClick: () => { },
                                                    'aria-label': 'Clear product filters',
                                                }}

                                            />
                                            <FilterPill
                                                name="Cloud instance"
                                                menuProps={{
                                                    items: filterCloudInstancesOptions, calloutProps: { calloutMaxHeight: 300 }
                                                }}
                                                filterPillMenuProps={{ headerText: 'Cloud instance' }}
                                                defaultValue={'All'}
                                                value={this._filter.cloudInstancesFilter}
                                                resetProps={{
                                                    onClick: () => { },
                                                    'aria-label': 'Clear cloud instance filters',
                                                }}
                                            />
                                            <FilterPill
                                                name="Platform"
                                                menuProps={{
                                                    items: filterPlatformsOptions, calloutProps: { calloutMaxHeight: 300 }
                                                }}
                                                filterPillMenuProps={{ headerText: 'Platform' }}
                                                defaultValue={'All'}
                                                value={this._filter.platformsFilter}
                                                resetProps={{
                                                    onClick: () => { },
                                                    'aria-label': 'Clear platform filter',
                                                }}
                                            />
                                            <FilterPill
                                                name="Release phase"
                                                menuProps={{
                                                    items: filterReleasePhaseOptions, calloutProps: { calloutMaxHeight: 300 }
                                                }}
                                                filterPillMenuProps={{ headerText: 'Release phase' }}
                                                defaultValue={'All'}
                                                value={this._filter.releasePhaseFilter}
                                                resetProps={{
                                                    onClick: () => { },
                                                    'aria-label': 'Clear release phase filter',
                                                }}
                                            />
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
                                        items={archivedItems}
                                        compact={false}
                                        columns={columns}
                                        selectionMode={SelectionMode.none}
                                        getKey={this._getKey}
                                        setKey="id"
                                        layoutMode={DetailsListLayoutMode.justified}
                                        isHeaderVisible={true}
                                    />
                                </PivotItem>
                            </Pivot>
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
                        headerText={selectedItemTitle}
                        isOpen={isDetailsPanelOpen}
                        onDismiss={this._onDismisDetailsPanel}
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
                                        disabled={this._isFirst(selectedItem?.id)}
                                        onClick={() => this._selectPrevious(selectedItem?.id)} />

                                    <IconButton
                                        iconProps={{ iconName: 'Down' }}
                                        title='Next'
                                        ariaLabel='Next item'
                                        disabled={this._isLast(selectedItem?.id)}
                                        onClick={() => this._selectNext(selectedItem?.id)} />
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
                        <RoadmapDetails id={selectedItem?.id} onExtendedPropertiesChange={this.handleExtendedPropertiesChange} onUpdateParentTitle={this._updatePanelTitle} />
                    </Panel>
                </div>
            </Fabric>
        );
    }

    componentDidMount() {
        var items: IRoadmapItem[] = [];
        var productList: IContextualMenuItem[] = [];
        var cloudInstanceList: IContextualMenuItem[] = [];
        var platformList: IContextualMenuItem[] = [];
        var releasePhaseList: IContextualMenuItem[] = [];
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

                fetch('/api/Roadmap', { headers: response.idToken === "" ? {} : { 'Authorization': `Bearer ${response.idToken}` } })
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
                            var links: IMessageLink[] = [];

                            if (message.moreInfoLink !== undefined)
                                links.push({
                                    name: "More information",
                                    url: message.moreInfoLink
                                });

                            if (message.link !== undefined)
                                links.push({
                                    name: "Roadmap",
                                    url: message.link
                                });

                            var categories;
                            var tags;
                            var cloudInstances;
                            var products;
                            var releasePhase;
                            var platforms;

                            if (!message.category || message.category.length <= 0) {
                                categories = [];
                            } else {
                                if (message.category[0].tagName === undefined) {
                                    categories = message.category;
                                } else {
                                    categories = message.category.map((t: IRoadmapTag) => t.tagName);
                                }
                            }

                            if (!message.tags || message.tags.length <= 0) {
                                tags = [];
                            } else {
                                if (message.tags[0].tagName === undefined) {
                                    tags = message.tags;
                                } else {
                                    tags = message.tags.map((t: IRoadmapTag) => t.tagName);
                                }
                            }

                            if (!message.cloudInstances || message.cloudInstances.length <= 0) {
                                cloudInstances = [];
                            } else {
                                if (message.cloudInstances[0].tagName === undefined) {
                                    cloudInstances = message.cloudInstances;
                                } else {
                                    cloudInstances = message.cloudInstances.map((t: IRoadmapTag) => t.tagName);
                                }
                            }

                            if (!message.products || message.products.length <= 0) {
                                products = [];
                            } else {
                                if (message.products[0].tagName === undefined) {
                                    products = message.products;
                                } else {
                                    products = message.products.map((t: IRoadmapTag) => t.tagName);
                                }
                            }

                            if (!message.releasePhase || message.releasePhase.length <= 0) {
                                releasePhase = [];
                            } else {
                                if (message.releasePhase[0].tagName === undefined) {
                                    releasePhase = message.releasePhase;
                                } else {
                                    releasePhase = message.releasePhase.map((t: IRoadmapTag) => t.tagName);
                                }
                            }

                            if (!message.platforms || message.platforms.length <= 0) {
                                platforms = [];
                            } else {
                                if (message.platforms[0].tagName === undefined) {
                                    platforms = message.platforms;
                                } else {
                                    platforms = message.platforms.map((t: IRoadmapTag) => t.tagName);
                                }
                            }

                            for (const product of products) {
                                if (!productList.find(s => s.text!.toLowerCase() === product.toLowerCase())) {
                                    productList.push({
                                        key: product.replace(" ", "").toLowerCase(),
                                        text: product,
                                        onClick: () => this._toggleProductFilterItem(product.replace(" ", "").toLowerCase()),
                                        canCheck: true
                                    });
                                }
                            }  

                            for (const cloudInstance of cloudInstances) {
                                if (!cloudInstanceList.find(s => s.text!.toLowerCase() === cloudInstance.toLowerCase())) {
                                    cloudInstanceList.push({
                                        key: cloudInstance.replace(" ", "").toLowerCase(),
                                        text: cloudInstance,
                                        onClick: () => this._toggleCloudInstanceFilterItem(cloudInstance.replace(" ", "").toLowerCase()),
                                        canCheck: true
                                    });
                                }
                            }

                            for (const rp of releasePhase) {
                                if (!releasePhaseList.find(s => s.text!.toLowerCase() === rp.toLowerCase())) {
                                    releasePhaseList.push({
                                        key: rp.replace(" ", "").toLowerCase(),
                                        text: rp,
                                        onClick: () => this._toggleReleasePhaseFilterItem(rp.replace(" ", "").toLowerCase()),
                                        canCheck: true
                                    });
                                }
                            }

                            for (const platform of platforms) {
                                if (!platformList.find(s => s.text!.toLowerCase() === platform.toLowerCase())) {
                                    platformList.push({
                                        key: platform.replace(" ", "").toLowerCase(),
                                        text: platform,
                                        onClick: () => this._togglePlatformFilterItem(platform.replace(" ", "").toLowerCase()),
                                        canCheck: true
                                    });
                                }
                            }

                            items.push({
                                id: message.id,
                                title: message.title,
                                links: links,
                                status: message.status ? message.status : 'not specified',
                                availabilityDate: message.availabilityDate,
                                availabilityFrom: new Date(message.availabilityFrom),
                                availabilityTo: new Date(message.availabilityTo),
                                publicPreviewDate: message.publicPreviewDate,
                                publicPreviewFrom: new Date(message.publicPreviewFrom),
                                publicPreviewTo: new Date(message.publicPreviewTo),
                                categories: categories,
                                tags: tags,
                                cloudInstances: cloudInstances,
                                products: products,
                                releasePhase: releasePhase,
                                platforms: platforms,
                                description: message.description,
                                published: new Date(message.published),
                                lastUpdated: new Date(message.lastUpdated),
                                extendedProperties: message.extendedProperties,
                                workItem: message.workItemId ? "<a href='" + message.workItemUrl + "' target='_blank'>" + message.workItemId + "</a>" : "",
                            });

                            items = _copyAndSort(items, 'lastUpdated', true);
                        }
                    }).then(() => {
                        productList = productList.sort((a, b) => 0 - (a.text! > b.text! ? -1 : 1));
                        cloudInstanceList = cloudInstanceList.sort((a, b) => 0 - (a.text! > b.text! ? -1 : 1));
                        releasePhaseList = releasePhaseList.sort((a, b) => 0 - (a.text! > b.text! ? -1 : 1));
                        platformList = platformList.sort((a, b) => 0 - (a.text! > b.text! ? -1 : 1));

                        this._allItems = items;
                        this.setState({
                            tags: this._getTags(this._allItems),
                            filterProductsOptions: productList,
                            filterCloudInstancesOptions: cloudInstanceList,
                            filterReleasePhaseOptions: releasePhaseList,
                            filterPlatformsOptions: platformList,
                            isDataLoaded: true
                        });
                        this._loadFilter();
                        this._filterItems();
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

    private _getTags(items: IRoadmapItem[]): ITag[] {
        var tags: ITag[] = [];

        for (const item of items) {
            for (const category of item.categories) {
                if (tags.filter(t => t.name === category).length <= 0)
                    tags.push({
                        key: category,
                        name: category
                    });
            }
        }

        return tags;
    }

    private _getKey(item: any, index?: number): string {
        return item.key;
    }

    _updatePanelTitle(title: string) {
        this.setState({
            selectedItemTitle: title
        });
    }

    private _selectPrevious(id: string | undefined): void {
        if (id !== undefined) {
            const itemIndex: number = this.state.items.findIndex((i) => i.id === id);

            if (itemIndex > 0) {
                /* this.state.selection.setIndexSelected(itemIndex, false, false);
                this.state.selection.setIndexSelected(itemIndex - 1, true, false); */
                this.setState({
                    selectedItem: this.state.items[itemIndex - 1],
                    selectedItemTitle: this.state.items[itemIndex - 1].title,
                    disablePanelNavigation: false
                });
            }
        }
    }

    private _selectNext(id: string | undefined): void {
        if (id !== undefined) {
            const itemIndex: number = this.state.items.findIndex((i) => i.id === id);

            if (itemIndex < this.state.items.length) {
                /* this.state.selection.setIndexSelected(itemIndex, false, false);
                this.state.selection.setIndexSelected(itemIndex + 1, true, false); */
                this.setState({
                    selectedItem: this.state.items[itemIndex + 1],
                    selectedItemTitle: this.state.items[itemIndex + 1].title,
                    disablePanelNavigation: false
                });
            }
        }
    }

    private _isFirst(id: string | undefined): boolean {
        if (this.state.items.length === 0)
            return true;

        const firstElement: IRoadmapItem = this.state.items[0];
        return id !== undefined ? firstElement.id === id : true;
    }

    private _isLast(id: string | undefined): boolean {
        if (this.state.items.length === 0)
            return true;
 
        const lastElement: IRoadmapItem = this.state.items[this.state.items.length - 1];
        return id !== undefined ? lastElement.id === id : true;
    }

    private _onChangeSearchText = (ev?: React.ChangeEvent<HTMLInputElement>, newValue?: string): void => {
        this._filter.textFilter = newValue ? newValue : "";
        this._filterItems();
    };

    private _filterItemsInt(archived: boolean): IRoadmapItem[] {
        var filteredItems: IRoadmapItem[];
        const textFilter: string = this._filter.textFilter;
        var productFilter: string[] = this._filter.productsFilter;
        var cloudInstanceFilter: string[] = this._filter.cloudInstancesFilter;
        var platformFilter: string[] = this._filter.platformsFilter;
        var releasePhaseFilter: string[] = this._filter.releasePhaseFilter;

        filteredItems = archived ?
            this._allItems.filter(i =>
                i.extendedProperties !== undefined &&
                i.extendedProperties !== null &&
                i.extendedProperties.Archived).slice() :
            this._allItems.filter(i =>
                !(i.extendedProperties !== undefined &&
                    i.extendedProperties !== null &&
                    i.extendedProperties.Archived)).slice();

        filteredItems = textFilter && textFilter !== "" ? filteredItems.filter(
            i => i.title.toLowerCase().indexOf(textFilter.toLowerCase()) > -1 ||
                i.id.toLowerCase().indexOf(textFilter.toLowerCase()) > -1 ||
                i.description.toLowerCase().indexOf(textFilter.toLowerCase()) > -1 ||
                i.categories.find(t => t?.toLowerCase().indexOf(textFilter.toLowerCase()) > -1)
        ) : filteredItems;

        if (productFilter.length > 0) {
            filteredItems = filteredItems.filter(i => productFilter.some(s => i.products.indexOf(s) >= 0));
        }

        if (cloudInstanceFilter.length > 0) {
            filteredItems = filteredItems.filter(i => cloudInstanceFilter.some(s => i.cloudInstances.indexOf(s) >= 0));
        }

        if (platformFilter.length > 0) {
            filteredItems = filteredItems.filter(i => platformFilter.some(s => i.platforms.indexOf(s) >= 0));
        }

        if (releasePhaseFilter.length > 0) {
            filteredItems = filteredItems.filter(i => releasePhaseFilter.some(s => i.releasePhase.indexOf(s) >= 0));
        }

        return filteredItems
    }

    private _filterItems(): void {
        var filteredInbox: IRoadmapItem[] = this._filterItemsInt(false);
        var filteredArchive: IRoadmapItem[] = this._filterItemsInt(true);

        this.setState({
            items: filteredInbox,
            archivedItems: filteredArchive
        });
    }

    private _toggleProductFilterItem(key: string): void {
        var productFilter: string[] = this._filter.productsFilter;
        var changed: boolean = false;
        var option: IContextualMenuItem | undefined = this.state.filterProductsOptions.find((o) => o.key === key);

        if (option !== undefined) {
            option.checked = !option.checked;

            if (option.checked) {
                if (productFilter.filter(s => s.toLowerCase() === option!.text!.toLowerCase()).length === 0) {
                    productFilter.push(option.text!);
                    changed = true;
                }
            } else {
                var idx: number = productFilter.indexOf(option!.text!);
                if (idx > -1) {
                    productFilter.splice(idx, 1);
                    changed = true;
                }
            }
        }

        if (changed) {
            this._filter.productsFilter = productFilter;
            this._filterItems();
        }
    }

    private _toggleCloudInstanceFilterItem(key: string): void {
        var cloudInstanceFilter: string[] = this._filter.cloudInstancesFilter;
        var changed: boolean = false;
        var option: IContextualMenuItem | undefined = this.state.filterCloudInstancesOptions.find((o) => o.key === key);

        if (option !== undefined) {
            option.checked = !option.checked;

            if (option.checked) {
                if (cloudInstanceFilter.filter(s => s.toLowerCase() === option!.text!.toLowerCase()).length === 0) {
                    cloudInstanceFilter.push(option.text!);
                    changed = true;
                }
            } else {
                var idx: number = cloudInstanceFilter.indexOf(option!.text!);
                if (idx > -1) {
                    cloudInstanceFilter.splice(idx, 1);
                    changed = true;
                }
            }
        }

        if (changed) {
            this._filter.cloudInstancesFilter = cloudInstanceFilter;
            this._filterItems();
        }
    }

    private _togglePlatformFilterItem(key: string): void {
        var platformFilter: string[] = this._filter.platformsFilter;
        var changed: boolean = false;
        var option: IContextualMenuItem | undefined = this.state.filterPlatformsOptions.find((o) => o.key === key);

        if (option !== undefined) {
            option.checked = !option.checked;

            if (option.checked) {
                if (platformFilter.filter(s => s.toLowerCase() === option!.text!.toLowerCase()).length === 0) {
                    platformFilter.push(option.text!);
                    changed = true;
                }
            } else {
                var idx: number = platformFilter.indexOf(option!.text!);
                if (idx > -1) {
                    platformFilter.splice(idx, 1);
                    changed = true;
                }
            }
        }

        if (changed) {
            this._filter.platformsFilter = platformFilter;
            this._filterItems();
        }
    }

    private _toggleReleasePhaseFilterItem(key: string): void {
        var releasePhaseFilter: string[] = this._filter.releasePhaseFilter;
        var changed: boolean = false;
        var option: IContextualMenuItem | undefined = this.state.filterReleasePhaseOptions.find((o) => o.key === key);

        if (option !== undefined) {
            option.checked = !option.checked;

            if (option.checked) {
                if (releasePhaseFilter.filter(s => s.toLowerCase() === option!.text!.toLowerCase()).length === 0) {
                    releasePhaseFilter.push(option.text!);
                    changed = true;
                }
            } else {
                var idx: number = releasePhaseFilter.indexOf(option!.text!);
                if (idx > -1) {
                    releasePhaseFilter.splice(idx, 1);
                    changed = true;
                }
            }
        }

        if (changed) {
            this._filter.releasePhaseFilter = releasePhaseFilter;
            this._filterItems();
        }
    }

    private _resetFilterOptions(options: IContextualMenuItem[]): void {
        for (const option of options)
            option.checked = false;
    }

    private _clearFilter(): void {
        this._filter.productsFilter = [];
        this._filter.cloudInstancesFilter = [];
        this._filter.platformsFilter = [];
        this._filter.releasePhaseFilter = [];

        this._resetFilterOptions(this.state.filterProductsOptions);
        this._resetFilterOptions(this.state.filterCloudInstancesOptions);
        this._resetFilterOptions(this.state.filterPlatformsOptions);
        this._resetFilterOptions(this.state.filterReleasePhaseOptions);

        this._filterItems();
    }

    private _setFilterOptions(filter: any, options: IContextualMenuItem[]) {
        for (const option of options)
            option.checked = filter.find((o: any) => o === option.text) !== undefined;
    }

    private _loadFilter(): void {
        let globalState: any = this.context;
        var roadmapFilter = globalState.getProfileProperty("roadmapFilter");
        if (roadmapFilter && roadmapFilter != null) {
            if (roadmapFilter.cloudInstancesFilter)
                this._filter.cloudInstancesFilter = roadmapFilter.cloudInstancesFilter;

            if (roadmapFilter.platformsFilter)
                this._filter.platformsFilter = roadmapFilter.platformsFilter;

            if (roadmapFilter.productsFilter)
                this._filter.productsFilter = roadmapFilter.productsFilter;

            if (roadmapFilter.releasePhaseFilter)
                this._filter.releasePhaseFilter = roadmapFilter.releasePhaseFilter;

            this._filter.textFilter = '';
        }

        this._setFilterOptions(this._filter.cloudInstancesFilter, this.state.filterCloudInstancesOptions);
        this._setFilterOptions(this._filter.platformsFilter, this.state.filterPlatformsOptions);
        this._setFilterOptions(this._filter.releasePhaseFilter, this.state.filterReleasePhaseOptions);
        this._setFilterOptions(this._filter.productsFilter, this.state.filterProductsOptions);
    }

    private _saveFilter(): void {
        let globalState: any = this.context;
        globalState.setProfileProperty("roadmapFilter", this._filter);
        globalState.setToastNotification({
            id: '0',
            title: 'Saved',
            message: 'Microsoft 365 Roadmap filter preferences saved.',
            time: new Date(),
            state: NotificationState.Information,
            active: true
        });
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

    _onOpenDetailsPanel(item: IRoadmapItem) {
        this.setState({
            selectedItem: item,
            selectedItemTitle: item.title,
            disablePanelNavigation: false,
            isDetailsPanelOpen: true
        });
    }

    private _onDismisDetailsPanel = (): void => {
        this.setState({
            selectedItem: undefined,
            disablePanelNavigation: false,
            isDetailsPanelOpen: false
        });
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
