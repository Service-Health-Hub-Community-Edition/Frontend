import * as React from 'react';
import { GlobalState } from '../../GlobalState';
import { NotificationState } from '../../Notifications';
import {
    IIconProps, ILinkStyleProps, ILinkStyles, Link, Text, SearchBox, FontIcon, IconButton, CommandBarButton,
    ResponsiveMode, Button, ITheme
} from '@fluentui/react';
import { Fabric, mergeStyles, mergeStyleSets } from '@fluentui/react';
import { DetailsList, DetailsListLayoutMode, Selection, SelectionMode, IColumn, IDetailsListStyles, ConstrainMode } from '@fluentui/react';
import { Panel, PanelType} from '@fluentui/react';
import { IDropdownOption, IContextualMenuItem } from '@fluentui/react';
import { CommandBar, ICommandBarItemProps, Pivot, PivotItem } from '@fluentui/react';
import { TooltipHost } from '@fluentui/react';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { IBreadcrumbItem } from '@fluentui/react';
import { M365Breadcrumb } from '@m365-admin/m365-breadcrumb';
import { DetailPageHeader } from '@m365-admin/detail-page';
import { MessageBar, MessageBarType, ThemeProvider } from '@fluentui/react';
import { HealthReport, HealthReportType, HealthReportMode, IHealthEvent } from '../dashboards/elements/Incidents';
import { acquireAccessToken } from "../../../auth/AccessTokenHelper";
import { checkInTeams } from '../../auth/detectTeams';
import { AccessDenied } from "../../AccessDenied";
import { InPageFilterSimple, FilterPill } from '@m365-admin/in-page-filter';
import { Filter, IFilter, FilterOperator } from "../../Filter";
import { v4 as uuidv4 } from 'uuid';

interface IM365Service {
    key: string;
    service: string;
    incidents: IHealthEvent[];
    advisories: IHealthEvent[];
}

export interface IServiceHealthState {
    services: IM365Service[];
    filter: Filter;
    filterChangeToken: string;
    itemCount: number;
    currentView: string | undefined;
    inTeams: boolean;
    error?: string;
    accessGranted?: boolean;
    theme?: any;
}

const searchIcon: IIconProps = { iconName: 'Search' };

const iconClass = mergeStyles({

});

const classNames = mergeStyleSets({
    incident: [{ color: 'red' }, iconClass],
    advisory: [{ color: 'rgb(0, 120, 212)' }, iconClass],
    healthy: [{ color: 'rgb(16, 124, 16)' }, iconClass]
});

export class ServiceHealth extends React.Component<{}, IServiceHealthState> {
    static contextType = GlobalState;
    private _breadcrumbItems: IBreadcrumbItem[] = []; 

    private _serviceHealthFilter: IFilter = {
        search: {
            property: ["id", "title", "summarySearchable"],
            value: undefined
        },
        service: {
            filter: [],
            property: "service",
            value: []
        },
        classification: {
            filter: [],
            property: "issueType",
            value: []
        },
        origins: {
            filter: [],
            property: "origin",
            value: []
        },
        status: {
            filter: [],
            property: "status",
            value: []
        },
        orgTags: {
            filter: [],
            property: "orgTags",
            value: undefined
        },
        published: {
            filter: [],
            property: "published",
            value: undefined
        },
        favorite: {
            filter: [],
            property: "favorite",
            value: undefined
        },
        lastModified: {
            filter: [],
            property: "lastModified",
            value: []
        }
    }

    private _serviceColumns: IColumn[] = [
        {
            key: 'clService',
            name: 'Service',
            fieldName: 'service',
            minWidth: 180,
            maxWidth: 400,
            isResizable: true,
            isCollapsible: false,
            isMultiline: false,
            data: 'string',
            onRender: (item: IM365Service) => {
                return (<Text><b>{item.service}</b></Text>);
            },
            isPadded: true,
        },
        {
            key: 'clStatus',
            name: 'Status',
            minWidth: 100,
            maxWidth: 300,
            isResizable: true,
            isCollapsible: true,
            isPadded: true,
            onRender: (item: IM365Service) => {
                return <>
                    {item.incidents.length > 0 ? (<>
                        <FontIcon iconName={'WarningSolid'} className={classNames.incident} style={{ marginRight: '4px' }} />{item.incidents.length} incident{item.incidents.length === 1 ? "s" : ""}
                        </>
                    ) : ""}&nbsp;
                    {item.advisories.length > 0 ? (<>
                        <FontIcon iconName={'InfoSolid'} className={classNames.advisory} style={{marginRight: '4px'}} />{item.advisories.length} advisor{item.advisories.length === 1 ? "y" : "ies"}
                    </>
                    ) : ""}
                    {item.incidents.length <= 0 && item.advisories.length <=0 ? (<>
                        <FontIcon iconName={'SkypeCircleCheck'} className={classNames.healthy} style={{ marginRight: '4px' }} />Healthy
                    </>
                    ) : ""}&nbsp;
                </>;
            },
        }
    ];
    constructor(props: {}) {
        super(props);

        this.state = {
            services: [],
            filter: new Filter(this._serviceHealthFilter, () => this._filterItems(), () => this._filterItems()),
            filterChangeToken: uuidv4(),
            itemCount: 0,
            currentView: "inbox",
            inTeams: checkInTeams(),
            error: undefined,
            accessGranted: undefined
        };

        this._breadcrumbItems = this.state.inTeams ? [ ] : 
            [
                { text: 'Home', key: 'home', isCurrentItem: false, href: '/' },
                { text: 'Service health', key: 'serviceHealth', isCurrentItem: true }
            ];
    }

    public render() {
        const {
            itemCount, currentView, services, error, accessGranted, theme, filter, filterChangeToken
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

        const _commandBarItems: ICommandBarItemProps[] = [
        ];

        const _commandBarFarItems: ICommandBarItemProps[] = [
            {
                key: 'itemsCount',
                onRender: () => {
                    return <div style={{ paddingTop: 4, paddingRight:8 }}><Text variant='smallPlus'><b>{itemCount} item{itemCount !== 1 ? "s" : ""}</b></Text></div>;
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
            <>
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
                                title="Service status"
                                description="View the health status of all services that are available with your current subscriptions. View info about the history of incidents and advisories that have been resolved."
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

                                <PivotItem headerText="Overview" id='inbox'>
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
                                                    menuProps={{ items: filter._filter.service.filter ? filter._filter.service.filter : [], calloutProps: { calloutMaxHeight: 300 } }}
                                                    filterPillMenuProps={{ headerText: 'Services' }}
                                                    defaultValue={'All'}
                                                    value={filter._filter.service.value}                                                    
                                                />
                                                <FilterPill
                                                    name="Classification"
                                                    menuProps={{
                                                        items: filter._filter.classification.filter ? filter._filter.classification.filter : [], calloutProps: { calloutMaxHeight: 300 }
                                                    }}
                                                    filterPillMenuProps={{ headerText: 'Classification' }}
                                                    defaultValue={'All'}
                                                    value={filter._filter.classification.value}
                                                />
                                                <FilterPill
                                                    name="Origin"
                                                    menuProps={{
                                                        items: filter._filter.origins.filter ? filter._filter.origins.filter : [], calloutProps: { calloutMaxHeight: 300 }
                                                    }}
                                                    filterPillMenuProps={{ headerText: 'Origin' }}
                                                    defaultValue={'All'}
                                                    value={filter._filter.origins.value}
                                                    resetProps={{
                                                        onClick: () => { },
                                                        'aria-label': 'Clear tags filters',
                                                    }}
                                                />
                                                <FilterPill
                                                    name="Status"
                                                    menuProps={{
                                                        items: filter._filter.status.filter ? filter._filter.status.filter : [], calloutProps: { calloutMaxHeight: 300 }
                                                    }}
                                                    filterPillMenuProps={{ headerText: 'Status' }}
                                                    defaultValue={'All'}
                                                    value={filter._filter.status.value}
                                                    resetProps={{
                                                        onClick: () => { },
                                                        'aria-label': 'Clear status filters',
                                                    }}
                                                />
                                                {filter._filter.orgTags.filter && filter._filter.orgTags.filter!.length > 0 ? (
                                                    <FilterPill
                                                        name="Org. tags"
                                                        menuProps={{
                                                            items: filter._filter.orgTags.filter ? filter._filter.orgTags.filter : [], calloutProps: { calloutMaxHeight: 300 }
                                                        }}
                                                        filterPillMenuProps={{ headerText: 'Org. tags' }}
                                                        defaultValue={'All'}
                                                        value={filter._filter.orgTags.value}
                                                        resetProps={{
                                                            onClick: () => { },
                                                            'aria-label': 'Clear organization tags filters',
                                                        }}
                                                    />) : (<></>)}
                                                <FilterPill
                                                    name="Last update"
                                                    menuProps={{
                                                        items: filter._filter.lastModified.filter ? filter._filter.lastModified.filter : [], calloutProps: { calloutMaxHeight: 300 }
                                                    }}
                                                    filterPillMenuProps={{ headerText: 'Last update' }}
                                                    defaultValue={'All'}
                                                    value={filter._filter.lastModified.value}
                                                    resetProps={{
                                                        onClick: () => { },
                                                        'aria-label': 'Clear last update filter',
                                                    }}
                                                />
                                                <FilterPill
                                                    name="Published"
                                                    menuProps={{
                                                        items: filter._filter.published.filter ? filter._filter.published.filter : [], calloutProps: { calloutMaxHeight: 300 }
                                                    }}
                                                    filterPillMenuProps={{ headerText: 'Published' }}
                                                    defaultValue={'All'}
                                                    value={filter._filter.published.value}
                                                    resetProps={{
                                                        onClick: () => { },
                                                        'aria-label': 'Clear last update filter',
                                                    }}
                                                />
                                                <FilterPill
                                                    name="Show"
                                                    menuProps={{
                                                        items: filter._filter.favorite.filter ? filter._filter.favorite.filter : [], calloutProps: { calloutMaxHeight: 300 }
                                                    }}
                                                    filterPillMenuProps={{ headerText: 'Items' }}
                                                    defaultValue={'All'}
                                                    value={filter._filter.favorite.value}
                                                    resetProps={{
                                                        onClick: () => { },
                                                        'aria-label': 'Clear last update filter',
                                                    }}
                                                />
                                                <CommandBarButton
                                                    iconProps={{ iconName: 'ClearFilter' }}
                                                    text="Reset all"
                                                    onClick={() => { this._clearFilter() }}
                                                />
                                            </InPageFilterSimple>
                                        </div>
                                    </div>
                                    <div className="row">
                                        <div className="col">
                                            <HealthReport
                                                type={HealthReportType.active}
                                                mode={HealthReportMode.page}
                                                filter={filter}
                                                filterChangeToken={filterChangeToken}
                                                onDataLoaded={(items: IHealthEvent[]) => this._loadFilterOptions(items)}
                                                onFilterChange={(numOfItems: number) => this._onFilterChange(numOfItems)}
                                                onReloadFilterOptions={(items: IHealthEvent[]) => this._reloadDynamicFilterOptions(items)}
                                            />
                                        </div>
                                    </div>
                                    <div className="row" style={{ paddingTop: '48px' }}>
                                        <div className="col">
                                            <Text variant="mediumPlus"><b>Microsoft service health</b></Text>
                                        </div>
                                    </div>
                                    <div className="row" style={{ paddingTop: '12px' }}>
                                        <div className="col">
                                            <Text variant="medium">Shows the current health status of your Microsoft services, and updates when the issues are fixed</Text>
                                        </div>
                                    </div>
                                    <div className="row">
                                        <div className="col">
                                            <DetailsList
                                                items={services}
                                                compact={false}
                                                columns={this._serviceColumns}
                                                selectionMode={SelectionMode.none}
                                                layoutMode={DetailsListLayoutMode.justified}
                                                constrainMode={ConstrainMode.horizontalConstrained}
                                                isHeaderVisible={true}
                                            />
                                        </div>
                                    </div>
                                </PivotItem>

                                <PivotItem headerText="History" id='archive'>
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
                                                    menuProps={{ items: filter._filter.service.filter ? filter._filter.service.filter : [], calloutProps: { calloutMaxHeight: 300 } }}
                                                    filterPillMenuProps={{ headerText: 'Services' }}
                                                    defaultValue={'All'}
                                                    value={filter._filter.service.value}
                                                />
                                                <FilterPill
                                                    name="Classification"
                                                    menuProps={{
                                                        items: filter._filter.classification.filter ? filter._filter.classification.filter : [], calloutProps: { calloutMaxHeight: 300 }
                                                    }}
                                                    filterPillMenuProps={{ headerText: 'Classification' }}
                                                    defaultValue={'All'}
                                                    value={filter._filter.classification.value}
                                                />
                                                <FilterPill
                                                    name="Origin"
                                                    menuProps={{
                                                        items: filter._filter.origins.filter ? filter._filter.origins.filter : [], calloutProps: { calloutMaxHeight: 300 }
                                                    }}
                                                    filterPillMenuProps={{ headerText: 'Tags' }}
                                                    defaultValue={'All'}
                                                    value={filter._filter.origins.value}
                                                    resetProps={{
                                                        onClick: () => { },
                                                        'aria-label': 'Clear tags filters',
                                                    }}
                                                />
                                                <FilterPill
                                                    name="Status"
                                                    menuProps={{
                                                        items: filter._filter.status.filter ? filter._filter.status.filter : [], calloutProps: { calloutMaxHeight: 300 }
                                                    }}
                                                    filterPillMenuProps={{ headerText: 'Status' }}
                                                    defaultValue={'All'}
                                                    value={filter._filter.status.value}
                                                    resetProps={{
                                                        onClick: () => { },
                                                        'aria-label': 'Clear tags filters',
                                                    }}
                                                />
                                                {filter._filter.orgTags.filter && filter._filter.orgTags.filter!.length > 0 ? (
                                                    <FilterPill
                                                        name="Org. tags"
                                                        menuProps={{
                                                            items: filter._filter.orgTags.filter ? filter._filter.orgTags.filter : [], calloutProps: { calloutMaxHeight: 300 }
                                                        }}
                                                        filterPillMenuProps={{ headerText: 'Org. tags' }}
                                                        defaultValue={'All'}
                                                        value={filter._filter.orgTags.value}
                                                        resetProps={{
                                                            onClick: () => { },
                                                            'aria-label': 'Clear organization tags filters',
                                                        }}
                                                    />) : (<></>)}
                                                <FilterPill
                                                    name="Last update"
                                                    menuProps={{
                                                        items: filter._filter.lastModified.filter ? filter._filter.lastModified.filter : [], calloutProps: { calloutMaxHeight: 300 }
                                                    }}
                                                    filterPillMenuProps={{ headerText: 'Last update' }}
                                                    defaultValue={'All'}
                                                    value={filter._filter.lastModified.value}
                                                    resetProps={{
                                                        onClick: () => { },
                                                        'aria-label': 'Clear last update filter',
                                                    }}
                                                />
                                                <FilterPill
                                                    name="Published"
                                                    menuProps={{
                                                        items: filter._filter.published.filter ? filter._filter.published.filter : [], calloutProps: { calloutMaxHeight: 300 }
                                                    }}
                                                    filterPillMenuProps={{ headerText: 'Published' }}
                                                    defaultValue={'All'}
                                                    value={filter._filter.published.value}
                                                    resetProps={{
                                                        onClick: () => { },
                                                        'aria-label': 'Clear last update filter',
                                                    }}
                                                />
                                                <FilterPill
                                                    name="Show"
                                                    menuProps={{
                                                        items: filter._filter.favorite.filter ? filter._filter.favorite.filter : [], calloutProps: { calloutMaxHeight: 300 }
                                                    }}
                                                    filterPillMenuProps={{ headerText: 'Items' }}
                                                    defaultValue={'All'}
                                                    value={filter._filter.favorite.value}
                                                    resetProps={{
                                                        onClick: () => { },
                                                        'aria-label': 'Clear last update filter',
                                                    }}
                                                />
                                                <CommandBarButton
                                                    iconProps={{ iconName: 'ClearFilter' }}
                                                    text="Reset all"
                                                    onClick={() => { this._clearFilter() }}
                                                />
                                            </InPageFilterSimple>

                                            <HealthReport
                                                type={HealthReportType.history}
                                                mode={HealthReportMode.page}
                                                filter={filter}
                                                filterChangeToken={filterChangeToken}
                                                onDataLoaded={(items: IHealthEvent[]) => this._loadFilterOptions(items)}
                                                onFilterChange={(numOfItems: number) => this._onFilterChange(numOfItems)}
                                                onReloadFilterOptions={(items: IHealthEvent[]) => this._reloadDynamicFilterOptions(items)}
                                            />
                                        </div>
                                    </div>
                                </PivotItem>
                            </Pivot>  
                        </div>
                    </div>
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
            </>
        );
    }

    public componentDidUpdate(previousProps: any, previousState: IServiceHealthState) {

    }

    componentDidMount() {
        var services: IM365Service[] = [];
        const requiredRoles: string[] = ['ServiceHealthReader', 'Communication.Write.All', 'Admin'];
        var userHasRequiredRole: boolean = false;

        let globalState: any = this.context;
        var theme: ITheme = globalState.getTheme();

        this.setState({
            theme: theme
        });

        acquireAccessToken()
            .then((response) => {
                var tokenClaims: any = response.account?.idTokenClaims;
                const userRoles: any = tokenClaims?.roles;
                userHasRequiredRole = userRoles.some((r: string) => requiredRoles.includes(r));

                this.setState({
                    accessGranted: userHasRequiredRole
                });

            }).catch((err) => {
                this.setState({
                    accessGranted: userHasRequiredRole,
                    error: err.message
                });
            });
    }

    private _loadServices(): IM365Service[] {
        var services: IM365Service[] = [];
        const requiredRoles: string[] = ['ServiceHealthReader', 'Communication.Write.All', 'Admin'];
        var userHasRequiredRole: boolean = false;

        acquireAccessToken()
            .then((response) => {
                var tokenClaims: any = response.account?.idTokenClaims;
                const userRoles: any = tokenClaims?.roles;
                userHasRequiredRole = userRoles.some((r: string) => requiredRoles.includes(r));

                if (userHasRequiredRole)
                    fetch('/api/microsoft365/healthOverview', { headers: response.idToken === "" ? {} : { 'Authorization': `Bearer ${response.idToken}` } })
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
                            for (const service of result) {
                                services.push({
                                    key: service.id,
                                    service: service.service,
                                    incidents: [],
                                    advisories: []
                                });
                            }
                        })
            }).catch((err) => {
                this.setState({
                    accessGranted: userHasRequiredRole,
                    error: err.message
                });
            });

        return services;
    }

    private _filterItems(): void {
        this.forceUpdate();
    }

    private _reloadDynamicFilterOptions(items: IHealthEvent[]): void {
        var orgTagList: IContextualMenuItem[] = [];

        for (const item of items) {
            for (const tag of item.orgTags) {
                if (!orgTagList.find(s => s.text?.toLowerCase() === tag.toLowerCase())) {
                    orgTagList.push({
                        key: tag.replace(" ", "").toLowerCase(),
                        text: tag,
                        onClick: () => this._setFilter('orgTags', tag.replace(" ", "").toLowerCase()),
                        canCheck: true
                    });
                }
            }
        }

        orgTagList = orgTagList.sort((a, b) => 0 - (a.text! > b.text! ? -1 : 1));

        var filter: Filter = this.state.filter;
        if (Array.isArray(filter._filter.orgTags.value)) {
            var newValues: string[] = JSON.parse(JSON.stringify(filter._filter.orgTags.value))
            for (const f of filter._filter.orgTags.value) {
                if (!orgTagList.find((v: IContextualMenuItem) => v.text?.toLowerCase().trim() === f.toLowerCase().trim())) {
                    newValues = newValues.filter((s: string) => s.toLowerCase().trim() !== f.toLowerCase().trim());
                }
            }

            filter._filter.orgTags.value = newValues;
        }

        filter._filter.orgTags.filter = orgTagList;

        this.setState({
            filter: filter,
            filterChangeToken: uuidv4()
        });
    }

    private _loadFilterOptions(items: IHealthEvent[]): void {
        var serviceList: IContextualMenuItem[] = [];
        var originList: IContextualMenuItem[] = [];
        var statusList: IContextualMenuItem[] = [];
        var classificationList: IContextualMenuItem[] = [];
        var orgTagList: IContextualMenuItem[] = [];
        var services: IM365Service[] = [];
        const requiredRoles: string[] = ['ServiceHealthReader', 'Communication.Write.All', 'Admin'];
        var userHasRequiredRole: boolean = false;

        for (const item of items) {
            for (const service of item.service) {
                if (!serviceList.find(s => s.text?.toLowerCase() === service.toLowerCase())) {
                    serviceList.push({
                        key: service.replace(" ", "").toLowerCase(),
                        text: service,
                        onClick: () => this._setFilter('service', service.replace(" ", "").toLowerCase()),
                        canCheck: true
                    });
                }
            }

            if (!classificationList.find(s => s.text?.toLowerCase() === item.issueType?.toLowerCase())) {
                classificationList.push({
                    key: item?.issueType.replace(" ", "").toLowerCase(),
                    text: item?.issueType,
                    onClick: () => this._setFilter('classification', item?.issueType.replace(" ", "").toLowerCase()),
                    canCheck: true
                });
            }

            if (!originList.find(s => s.text?.toLowerCase() === item.origin?.toLowerCase())) {
                originList.push({
                    key: item?.origin.replace(" ", "").toLowerCase(),
                    text: item?.origin,
                    onClick: () => this._setFilter('origins', item?.origin.replace(" ", "").toLowerCase()),
                    canCheck: true
                });
            }

            if (!statusList.find(s => s.text?.toLowerCase() === item.status?.toLowerCase())) {
                statusList.push({
                    key: item?.status.replace(" ", "").toLowerCase(),
                    text: item?.status,
                    onClick: () => this._setFilter('status', item?.status.replace(" ", "").toLowerCase()),
                    canCheck: true
                });
            }

            for (const tag of item.orgTags) {
                if (!orgTagList.find(s => s.text?.toLowerCase() === tag.toLowerCase())) {
                    orgTagList.push({
                        key: tag.replace(" ", "").toLowerCase(),
                        text: tag,
                        onClick: () => this._setFilter('orgTags', tag.replace(" ", "").toLowerCase()),
                        canCheck: true
                    });
                }
            }
        }

        serviceList = serviceList.sort((a, b) => 0 - (a.text! > b.text! ? -1 : 1));
        originList = originList.sort((a, b) => 0 - (a.text! > b.text! ? -1 : 1));
        statusList = statusList.sort((a, b) => 0 - (a.text! > b.text! ? -1 : 1));
        classificationList = classificationList.sort((a, b) => 0 - (a.text! > b.text! ? -1 : 1));
        orgTagList = orgTagList.sort((a, b) => 0 - (a.text! > b.text! ? -1 : 1));

        var filter: IFilter = {
            search: {
                property: ["id", "title", "summarySearchable"],
                value: undefined
            },
            service: {
                filter: serviceList,
                property: "service",
                value: []
            },
            classification: {
                filter: classificationList,
                property: "issueType",
                value: []
            },
            origins: {
                filter: originList,
                property: "origin",
                value: []
            },
            status: {
                filter: statusList,
                property: "status",
                value: []
            },
            orgTags: {
                filter: orgTagList,
                property: "orgTags",
                value: []
            },
            favorite: {
                filter: [
                    { key: 'favorite', text: 'Favorites only', onClick: () => this._setFilter('favorite', 'favorite', true), canCheck: true, checked: false }
                ],
                property: "favorite",
                value: undefined,
                convertValue: (value) => {
                    const convertToBoolean = (value: any): boolean | undefined => {
                        switch (value) {
                            case "Favorites only":
                                return true;
                            default:
                                return undefined;
                        }
                    }

                    if (value === undefined)
                        return undefined
                    else {
                        var val: any = Array.isArray(value) ? (value.length > 0 ? value[0] : "") : value;
                        return convertToBoolean(val);
                    }
                }
            },
            published: {
                filter: [
                    { key: 'published', text: 'Yes', onClick: () => this._setFilter('published', 'published', true), canCheck: true, checked: false },
                    { key: 'notpublished', text: 'No', onClick: () => this._setFilter('published', 'notpublished', true), canCheck: true, checked: false }
                ],
                property: "published",
                value: undefined,
                convertValue: (value) => {
                    const convertToBoolean = (value: any): boolean | undefined => {
                        switch (value) {
                            case "Yes":
                                return true;
                            case "No":
                                return false;
                            default:
                                return undefined;
                        }
                    }

                    if (value === undefined)
                        return undefined
                    else {
                        var val: any = Array.isArray(value) ? (value.length > 0 ? value[0] : "") : value;
                        return convertToBoolean(val);
                    }
                }
            },
            lastModified: {
                filter: [
                    { key: '7', text: '7 days', onClick: () => this._setFilter('lastModified', '7', true), canCheck: true, checked: false },
                    { key: '30', text: '30 days', onClick: () => this._setFilter('lastModified', '30', true), canCheck: true, checked: false },
                    { key: '90', text: '90 days', onClick: () => this._setFilter('lastModified', '90', true), canCheck: true, checked: false },
                    { key: '180', text: '180 days', onClick: () => this._setFilter('lastModified', '180', true), canCheck: true, checked: false }
                ],
                property: "lastModified",
                operator: FilterOperator.ge,
                value: [],
                convertValue: (value: any, operator?: FilterOperator) => {
                    const convertRangeToDate = (dateRange: string, operator?: FilterOperator): Date => {
                        const getDate = (value: number, type: string, operator: FilterOperator): Date => {
                            var d: Date = new Date();
                            switch (type) {
                                case "day":
                                    return operator == FilterOperator.ge || FilterOperator.gt ?
                                        new Date(d.setDate(d.getDate() - value)) :
                                        new Date(d.setDate(d.getDate() + value));
                                case "month":
                                    return operator == FilterOperator.ge || FilterOperator.gt ?
                                        new Date(d.setMonth(d.getMonth() - value)) :
                                        new Date(d.setMonth(d.getMonth() + value));
                                default:
                                    return operator == FilterOperator.ge || FilterOperator.gt ?
                                        new Date(d.setDate(d.getDate() - value)) :
                                        new Date(d.setDate(d.getDate() + value));
                            }
                        }

                        const op: FilterOperator = operator ? operator : FilterOperator.ge;
                        const dateRangeFragments: string[] = dateRange.split(" ");
                        var dateElementType: string = dateRangeFragments.length < 2 ? "days" : dateRangeFragments[1];
                        switch (dateElementType.toLowerCase()) {
                            case "day":
                            case "days":
                                dateElementType = "day";
                                break;
                            case "month":
                            case "months":
                                dateElementType = "month";
                                break;
                            default:
                                dateElementType = "day";
                        }

                        var dateElementValue: number = dateRangeFragments.length < 1 ? 0 : +dateRangeFragments[0];
                        if (Number.isNaN(dateElementValue))
                            dateElementValue = 0;

                        return getDate(dateElementValue, dateElementType, op);
                    }

                    if (value === undefined)
                        return undefined;

                    if (Array.isArray(value)) {
                        var returnValue: Date[] = [];
                        for (const val of value)
                            returnValue.push(convertRangeToDate(val, operator));
                        return returnValue;
                    } else {
                        return convertRangeToDate(value, operator);
                    }
                }
            }
        }

        if (this.state.filter) {
            this.state.filter.setFilter(filter);
            this.setState({
                itemCount: items.length,
                filterChangeToken: uuidv4()
            });
        }
        else
            this.setState({
                itemCount: items.length,
                filter: new Filter(filter, () => this._filterItems(), () => this._filterItems()),
                filterChangeToken: uuidv4()
            });

        if (this.state.currentView === "archive")
            this._setFilter('lastModified', '7', true);

        acquireAccessToken()
            .then((response) => {
                var tokenClaims: any = response.account?.idTokenClaims;
                const userRoles: any = tokenClaims?.roles;
                userHasRequiredRole = userRoles.some((r: string) => requiredRoles.includes(r));

                if (userHasRequiredRole)
                    fetch('/api/microsoft365/healthOverview', { headers: response.idToken === "" ? {} : { 'Authorization': `Bearer ${response.idToken}` } })
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
                            for (const service of result) {
                                services.push({
                                    key: service.id,
                                    service: service.service,
                                    incidents: [],
                                    advisories: []
                                });
                            }

                            for (const svc of services) {
                                svc.incidents = [];
                                svc.advisories = [];
                            }

                            for (const item of items) {
                                for (const svc of item.service) {
                                    var m365Service: IM365Service | undefined = services.find(s => s.service === svc);
                                    if (m365Service) {
                                        switch (item.issueType) {
                                            case "Incident":
                                                m365Service.incidents.push(item);
                                                break;
                                            case "Advisory":
                                                m365Service.advisories.push(item);
                                                break;
                                            default:
                                                break;
                                        }
                                    }
                                }
                            }

                            services = services.sort((a, b) => 0 - (a.service! > b.service! ? -1 : 1));
                            services = services.sort((a, b) => ((a.incidents.length > 0 || a.advisories.length > 0) > (b.incidents.length > 0 || b.advisories.length > 0) ? -1 : 1));

                            this.setState({
                                services: services
                            });
                        })
            }).catch((err) => {
                this.setState({
                    accessGranted: userHasRequiredRole,
                    error: err.message
                });
            });
    }

    private _onChangeSearchText = (ev?: React.ChangeEvent<HTMLInputElement>, newValue?: string): void => {
        this.state.filter.setFilterSearchItem("search", newValue);
        this.setState({
            filterChangeToken: uuidv4()
        });
    };

    private _setFilter(filterKey: string, key: string, singleItemSelection?: boolean): void {
        this.state.filter?.setFilterItem(filterKey, key, singleItemSelection);
        this.setState({
            filterChangeToken: uuidv4()
        });
    }

    private _clearFilter(): void {
        this.state.filter?.clearFilter();
        this.setState({
            filterChangeToken: uuidv4()
        });
    }

    private _onFilterChange(itemCount: number): void {
        this.setState({
            itemCount: itemCount
        });
    }

    /* private _setFilterOptions(filter: any, options: IContextualMenuItem[]) {
        for (const option of options)
            option.checked = filter.find((o: any) => o === option.text) !== undefined;
    }

    private _loadFilter(): void {
        let globalState: any = this.context;
        var messageCenterFilter = globalState.getProfileProperty("serviceHealthFilter");
        if (messageCenterFilter && messageCenterFilter != null) {
            if (messageCenterFilter.lastUpdateFilter)
                this._filter.lastUpdateFilter = messageCenterFilter.lastUpdateFilter;

            if (messageCenterFilter.publishedFilter)
                this._filter.publishedFilter = messageCenterFilter.publishedFilter;

            if (messageCenterFilter.serviceFilter)
                this._filter.serviceFilter = messageCenterFilter.serviceFilter;

            if (messageCenterFilter.tagFilter)
                this._filter.tagFilter = messageCenterFilter.tagFilter;

            this._filter.textFilter = '';
        }

        this._setFilterOptions(this._filter.lastUpdateFilter, this.state.filterLastUpdateOptions);
        this._setFilterOptions(this._filter.publishedFilter, this.state.filterPublishedOptions);
        this._setFilterOptions(this._filter.serviceFilter, this.state.filterServicesOptions);
        this._setFilterOptions(this._filter.tagFilter, this.state.filterTagOptions);
    } */

    private _saveFilter(): void {
        /*let globalState: any = this.context;
        globalState.setProfileProperty("serviceHealthFilter", this._filter);
        globalState.setToastNotification({
            id: '0',
            title: 'Saved',
            message: 'Service health filter preferences saved.',
            time: new Date(),
            state: NotificationState.Information,
            active: true
        });*/
    }

    _onArchivePivotChange = (item?: PivotItem, ev?: React.MouseEvent<HTMLElement>): void => {
        if (item)
            this.setState({
                itemCount: 0,
                currentView: item.props.id
            });
    }
}