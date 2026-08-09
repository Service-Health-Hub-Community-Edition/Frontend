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
import { AzureUpdatesReport, AzureUpdatesReportType, AzureUpdatesReportMode, IAzureUpdateEvent } from '../dashboards/elements/Updates';
import { acquireAccessToken } from "../../../auth/AccessTokenHelper";
import { checkInTeams } from '../../auth/detectTeams';
import { AccessDenied } from "../../AccessDenied";
import { InPageFilterSimple, FilterPill } from '@m365-admin/in-page-filter';
import { Filter, IFilter, FilterOperator } from "../../Filter";
import { v4 as uuidv4 } from 'uuid';

export interface IAzureUpdatesState {
    filter: Filter;
    filterChangeToken: string;
    itemCount: number;
    currentView: string | undefined;
    inTeams: boolean;
    error?: string;
    accessGranted?: boolean;
    theme?: any;
}

export interface IAzureUpdateMessage {
    key: string;
    name: string;
    value: string;
    id: string;
    title: string;
    service: string[];
    origin: string;
    lastUpdated: Date;
    actBy?: Date;
    tags: string[];
    extendedProperties: any;
    published: boolean;
}

const searchIcon: IIconProps = { iconName: 'Search' };

const iconClass = mergeStyles({

});

const classNames = mergeStyleSets({
    incident: [{ color: 'red' }, iconClass],
    advisory: [{ color: 'rgb(0, 120, 212)' }, iconClass],
    healthy: [{ color: 'rgb(16, 124, 16)' }, iconClass]
});

export class AzureUpdates extends React.Component<{}, IAzureUpdatesState> {
    static contextType = GlobalState;
    private _breadcrumbItems: IBreadcrumbItem[] = []; 

    private _azureUpdatesFilter: IFilter = {
        search: {
            property: ["id", "title"],
            value: undefined
        },
        tags: {
            filter: [],
            property: "tags",
            value: []
        },
        releaseStatus: {
            filter: [],
            property: "releaseStatus",
            value: []
        },
        orgTags: {
            filter: [],
            property: "orgTags",
            value: undefined
        },
        favorite: {
            filter: [],
            property: "favorite",
            value: undefined
        },
        published: {
            filter: [
                { key: '7', text: '7 days', onClick: () => this._setFilter('published', '7', true), canCheck: true, checked: false },
                { key: '30', text: '30 days', onClick: () => this._setFilter('published', '30', true), canCheck: true, checked: false },
                { key: '90', text: '90 days', onClick: () => this._setFilter('published', '90', true), canCheck: true, checked: false },
                { key: '180', text: '180 days', onClick: () => this._setFilter('published', '180', true), canCheck: true, checked: false }
            ],
            property: "published",
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

    constructor(props: {}) {
        super(props);

        this.state = {
            filter: new Filter(this._azureUpdatesFilter, () => this._filterItems(), () => this._filterItems()),
            filterChangeToken: uuidv4(),
            itemCount: 0,
            currentView: "inbox",
            inTeams: checkInTeams(),
            error: undefined,
            accessGranted: undefined
        };

        this._breadcrumbItems = this.state.inTeams ? [ ] : 
            [
                { text: 'Home', key: 'home', isCurrentItem: false, href: '/azure' },
                { text: 'Azure updates', key: 'azureUpdates', isCurrentItem: true }
            ];
    }

    public render() {
        const {
            itemCount, currentView, error, accessGranted, theme, filter, filterChangeToken
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
                                title="Azure updates"
                                description="View the list of upcoming Azure updates. This overview contains general information about global Azure announcements, upcoming features and functionality changes."
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

                                <PivotItem headerText="Active" id='inbox'>
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
                                                    name="Tags"
                                                    menuProps={{ items: filter._filter.tags.filter ? filter._filter.tags.filter : [], calloutProps: { calloutMaxHeight: 300 } }}
                                                    filterPillMenuProps={{ headerText: 'Tags' }}
                                                    defaultValue={'All'}
                                                    value={filter._filter.tags.value}                                                    
                                                />
                                                <FilterPill
                                                    name="Release status"
                                                    menuProps={{
                                                        items: filter._filter.releaseStatus.filter ? filter._filter.releaseStatus.filter : [], calloutProps: { calloutMaxHeight: 300 }
                                                    }}
                                                    filterPillMenuProps={{ headerText: 'Release status' }}
                                                    defaultValue={'All'}
                                                    value={filter._filter.releaseStatus.value}
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
                                                    name="Published"
                                                    menuProps={{
                                                        items: filter._filter.published.filter ? filter._filter.published.filter : [], calloutProps: { calloutMaxHeight: 300 }
                                                    }}
                                                    filterPillMenuProps={{ headerText: 'Published' }}
                                                    defaultValue={'All'}
                                                    value={filter._filter.published.value}
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
                                            <AzureUpdatesReport
                                                type={AzureUpdatesReportType.active}
                                                mode={AzureUpdatesReportMode.page}
                                                filter={filter}
                                                filterChangeToken={filterChangeToken}
                                                onDataLoaded={(items: IAzureUpdateEvent[]) => this._loadFilterOptions(items)}
                                                onFilterChange={(numOfItems: number) => this._onFilterChange(numOfItems)}
                                                onReloadFilterOptions={(items: IAzureUpdateEvent[]) => this._reloadDynamicFilterOptions(items)}
                                            />
                                        </div>
                                    </div>
                                </PivotItem>

                                <PivotItem headerText="Released" id='archive'>
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
                                                    name="Tags"
                                                    menuProps={{ items: filter._filter.tags.filter ? filter._filter.tags.filter : [], calloutProps: { calloutMaxHeight: 300 } }}
                                                    filterPillMenuProps={{ headerText: 'Tags' }}
                                                    defaultValue={'All'}
                                                    value={filter._filter.tags.value}
                                                />
                                                <FilterPill
                                                    name="Release status"
                                                    menuProps={{
                                                        items: filter._filter.releaseStatus.filter ? filter._filter.releaseStatus.filter : [], calloutProps: { calloutMaxHeight: 300 }
                                                    }}
                                                    filterPillMenuProps={{ headerText: 'Release status' }}
                                                    defaultValue={'All'}
                                                    value={filter._filter.releaseStatus.value}
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
                                                    name="Published"
                                                    menuProps={{
                                                        items: filter._filter.published.filter ? filter._filter.published.filter : [], calloutProps: { calloutMaxHeight: 300 }
                                                    }}
                                                    filterPillMenuProps={{ headerText: 'Published' }}
                                                    defaultValue={'All'}
                                                    value={filter._filter.published.value}
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

                                            <AzureUpdatesReport
                                                type={AzureUpdatesReportType.history}
                                                mode={AzureUpdatesReportMode.page}
                                                filter={filter}
                                                filterChangeToken={filterChangeToken}
                                                onDataLoaded={(items: IAzureUpdateEvent[]) => this._loadFilterOptions(items)}
                                                onFilterChange={(numOfItems: number) => this._onFilterChange(numOfItems)}
                                                onReloadFilterOptions={(items: IAzureUpdateEvent[]) => this._reloadDynamicFilterOptions(items)}
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

    public componentDidUpdate(previousProps: any, previousState: IAzureUpdatesState) {

    }

    componentDidMount() {
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

    private _filterItems(): void {
        this.forceUpdate();
    }

    private _reloadDynamicFilterOptions(items: IAzureUpdateEvent[]): void {
        var orgTagList: IContextualMenuItem[] = [];

        for (const item of items) {
            for (const tag of item.orgTags) {
                if (tag !== null && !orgTagList.find(s => s.text?.toLowerCase() === tag.toLowerCase())) {
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

    private _loadFilterOptions(items: IAzureUpdateEvent[]): void {
        var tagList: IContextualMenuItem[] = [];
        var releaseStatusList: IContextualMenuItem[] = [];
        var orgTagList: IContextualMenuItem[] = [];

        for (const item of items) {
            for (const tag of item.tags) {
                console.log(tag)
                if (tag !== null && !tagList.find(t => t.text?.toLowerCase() === tag.toLowerCase())) {
                    tagList.push({
                        key: tag.replace(" ", "").toLowerCase(),
                        text: tag,
                        onClick: () => this._setFilter('tags', tag.replace(" ", "").toLowerCase()),
                        canCheck: true
                    });
                }
            }

            if (item.releaseStatus !== null && !releaseStatusList.find(s => s.text?.toLowerCase() === item.releaseStatus?.toLowerCase())) {
                releaseStatusList.push({
                    key: item?.releaseStatus.replace(" ", "").toLowerCase(),
                    text: item?.releaseStatus,
                    onClick: () => this._setFilter('type', item?.releaseStatus.replace(" ", "").toLowerCase()),
                    canCheck: true
                });
            }

            for (const tag of item.orgTags) {
                if (tag !== null && !orgTagList.find(s => s.text?.toLowerCase() === tag.toLowerCase())) {
                    orgTagList.push({
                        key: tag.replace(" ", "").toLowerCase(),
                        text: tag,
                        onClick: () => this._setFilter('orgTags', tag.replace(" ", "").toLowerCase()),
                        canCheck: true
                    });
                }
            }
        }

        tagList = tagList.sort((a, b) => 0 - (a.text! > b.text! ? -1 : 1));
        releaseStatusList = releaseStatusList.sort((a, b) => 0 - (a.text! > b.text! ? -1 : 1));
        orgTagList = orgTagList.sort((a, b) => 0 - (a.text! > b.text! ? -1 : 1));

        var filter: IFilter = {
            search: {
                property: ["id", "title", "summary"],
                value: undefined
            },
            tags: {
                filter: tagList,
                property: "tags",
                value: []
            },
            releaseStatus: {
                filter: releaseStatusList,
                property: "releaseStatus",
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
                    { key: '7', text: '7 days', onClick: () => this._setFilter('published', '7', true), canCheck: true, checked: false },
                    { key: '30', text: '30 days', onClick: () => this._setFilter('published', '30', true), canCheck: true, checked: false },
                    { key: '90', text: '90 days', onClick: () => this._setFilter('published', '90', true), canCheck: true, checked: false },
                    { key: '180', text: '180 days', onClick: () => this._setFilter('published', '180', true), canCheck: true, checked: false }
                ],
                property: "published",
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

    _onArchivePivotChange = (item?: PivotItem, ev?: React.MouseEvent<HTMLElement>): void => {
        if (item)
            this.setState({
                itemCount: 0,
                currentView: item.props.id
            });
    }
}