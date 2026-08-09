import * as React from 'react';
import { ColumnActionsMode, DefaultPalette } from '@fluentui/react';
import { Text, Link, ILinkStyleProps, ILinkStyles, TooltipHost, ScrollablePane, Sticky, IDetailsHeaderProps, IRenderFunction } from '@fluentui/react';
import {
    IColumn, DetailsList, SelectionMode, DetailsListLayoutMode, ConstrainMode, FontIcon,
    IDetailsListStyles, mergeStyles, mergeStyleSets, IGroup,
    HoverCard, IExpandingCardProps, DirectionalHint, IconButton, IContextualMenuProps
} from '@fluentui/react';
import { Panel, PanelType } from '@fluentui/react';
import { Tag, TagType, ITagProps } from '@m365-admin/tag';
import { IncidentDetails } from '../../../IncidentDetails';
import { ICustomAction, CustomAction } from '../../../CustomAction';
import { ServiceComponent } from '../../../ServiceNameComponent';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { MessageBar, MessageBarType } from '@fluentui/react';
import { acquireAccessToken, accessControl } from "../../../../auth/AccessTokenHelper";
import { AuthenticationResult } from '@azure/msal-browser';
import { checkInTeams } from '../../../auth/detectTeams';
import { AccessDenied } from "../../../AccessDenied";
import { Filter } from "../../../Filter";
import { setViewState, setArchiveState, setFavoriteState } from "../../../../api/viewpoint";
import { translate, getUIString } from "../../../../translation/translator";
import { IPublicEvents, IPublicIncident } from '../../../IServiceStatistics';
import { GlobalState } from "../../../GlobalState";
import { PublicIncidentDetails } from "./PublicIncidentDetails"

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
}

interface IHealthReportState {
    filter?: Filter;
    filterChangeToken?: string;
    statistics?: IPublicEvents;
    items: IPublicIncident[];
    incidentGroups: IGroup[];
    isDataLoaded: boolean;
    isDetailsPanelOpen: boolean;
    selectedIncident: string;
    uiStrings: Map<string, string>;
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

export class PublicHealthReport extends React.Component<{
    type: HealthReportType,
    mode?: HealthReportMode,
    collectionMode?: HealthItems,
    filter?: Filter,
    filterChangeToken?: string,
    statistics: IPublicEvents | undefined,
    onDataLoaded?: (items: IPublicIncident[]) => void,
    onFilterChange?: (itemCount: number) => void
}, IHealthReportState> {
    static contextType = GlobalState;
    private _allItems: IPublicIncident[] = [];

    constructor(props: {
        type: HealthReportType,
        mode?: HealthReportMode,
        collectionMode?: HealthItems,
        filter?: Filter,
        filterChangeToken?: string,
        statistics: IPublicEvents | undefined,
        onDataLoaded?: (items: IPublicIncident[]) => void,
        onFilterChange?: (itemCount: number) => void
    }) {
        super(props);

        var uiStrings: Map<string, string> = new Map([
            ["Title", "Title"],
            ["Active issues", "Active issues"],
            ["Affected services", "Affected services"],
            ["Start", "Start"],
            ["Last update", "Last update"],
            ["Couldn't retrieve data. Error:", "Couldn't retrieve data. Error:"],
            ["Close", "Close"],
            ["There are no active service events!", "There are no active service events!"]
        ]);

        this.state = {
            filter: props.filter,
            filterChangeToken: props.filterChangeToken,
            statistics: this.props.statistics,
            items: [],
            isDetailsPanelOpen: false,
            selectedIncident: "",
            isDataLoaded: false,
            uiStrings: uiStrings,
            incidentGroups: [],
            inTeams: checkInTeams(),
            error: undefined,
            accessGranted: undefined
        };
    }

    public render() {
        const { items, incidentGroups, isDetailsPanelOpen, selectedIncident, isDataLoaded, uiStrings, error } = this.state;

        const pipeFabricStyles = (p: ILinkStyleProps): ILinkStyles => ({
            root: {
                textDecoration: 'none',
                color: p.theme.semanticColors.bodyText
            },
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
                onRender: (item: IPublicIncident) => {
                    return <div className="container" style={{ cursor: 'default' }} >
                            <div className="row" >
                                <div className="col" style={{ justifyContent: 'center', paddingLeft: '12px', paddingRight: '0px' }} >
                                    <FontIcon
                                        iconName={'WarningSolid'}
                                        className={classNames.incident} />

                                </div>
                            </div></div>;
                },
                isPadded: true,
            },
            {
                key: 'clTitle',
                name: getUIString(uiStrings, 'Title'),
                fieldName: 'title',
                minWidth: 180,
                isResizable: true,
                isCollapsible: false,
                isMultiline: false,
                data: 'string',
                onRender: (item: IPublicIncident) => {
                    return (<Link onClick={(event) => {
                        event.preventDefault();
                        this._onOpenDetailsPanel(item.id);
                    }} styles={pipeFabricStyles}>{
                                item.title 
                        }</Link >);
                },
                isPadded: false,
            },
            {
                key: 'clStart',
                name: getUIString(uiStrings, 'Start'),
                minWidth: 120,
                maxWidth: 120,
                isResizable: false,
                isCollapsible: true,
                isPadded: false,
                onRender: (item: IPublicIncident) => {
                    return <>{item.startTime ?
                        new Intl.DateTimeFormat(
                            window.navigator.language,
                            {
                                day: 'numeric',
                                month: 'numeric',
                                year: 'numeric',
                                hour: 'numeric',
                                minute: 'numeric'
                            }).format(new Date(item.startTime)) : ""}</>;
                },
            },
            {
                key: 'clUpdated',
                name: getUIString(uiStrings, 'Last update'),
                minWidth: 120,
                maxWidth: 120,
                isResizable: false,
                isCollapsible: true,
                isPadded: false,
                onRender: (item: IPublicIncident) => {
                    return <>{item.lastModified ?
                        new Intl.DateTimeFormat(
                            window.navigator.language,
                            {
                                day: 'numeric',
                                month: 'numeric',
                                year: 'numeric',
                                hour: 'numeric',
                                minute: 'numeric'
                            }).format(new Date(item.lastModified)) : ""}</>;
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
                onRender: (item: IPublicIncident) => {
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
            {this.props.mode !== HealthReportMode.dashboard ? (
                <DetailsList
                    items={items}
                    compact={true}
                    columns={columns}
                    groups={incidentGroups}
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
                            groups={incidentGroups}
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
                headerText={items.find(m => m.id.toLowerCase() === selectedIncident.toLowerCase())?.title!}
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
                                disabled={this._isFirst(selectedIncident)}
                                onClick={() => this._selectPrevious(selectedIncident)} />

                            <IconButton
                                iconProps={{ iconName: 'Down' }}
                                title='Next'
                                ariaLabel='Next item'
                                disabled={this._isLast(selectedIncident)}
                                onClick={() => this._selectNext(selectedIncident)} />

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
                <PublicIncidentDetails
                    incidentDetails={items.find(m => m.id.toLowerCase() === selectedIncident.toLowerCase())}
                /> 
            </Panel>
        </>
        );
    }

    componentDidMount() {
        this._getIncidents();
    }

    private _getIncidents() {
        let globalState: any = this.context;
        var language: string = globalState.getProfileProperty("language", "en").trim();
        var translatedUIStrings: Map<string, string>;
        var langQuery: string = "";
        if (language !== "" && language !== "en") {
            langQuery = "?lang=" + language;
        }

        var authResponse: AuthenticationResult;

        acquireAccessToken()
            .then((response) => {
                authResponse = response;

                translate(language,
                    this.state.uiStrings,
                    authResponse
                ).then((response) => {
                    translatedUIStrings = response;
                    this.setState({
                        uiStrings: translatedUIStrings
                    });
                });
            }).then(() => {
                if (this.state.statistics === undefined) {
                    fetch('/api/PublicReport' + langQuery, { headers: authResponse.idToken === "" ? {} : { 'Authorization': `Bearer ${authResponse.idToken}` } })
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
                            this._allItems = result.events ? result.events : [];
                            try {
                                if (this.props.onDataLoaded)
                                    this.props.onDataLoaded(this._allItems)
                            }
                            catch (err) {
                                this.setState({
                                    error: err.message
                                });
                            }
                            this._processIncidents(translatedUIStrings);
                        }).catch((err) => {
                            this.setState({
                                error: err.message
                            });
                        });
                }
                else {
                    this._allItems = this.state.statistics?.events ? this.state.statistics?.events : [];
                    this._processIncidents(translatedUIStrings);
                }
            });
    }

    componentDidUpdate(prevProps: any) {
        if (this.props.filterChangeToken && this.props.filterChangeToken !== this.state.filterChangeToken) {
            this._processIncidents(this.state.uiStrings);

            if (this.props.onFilterChange)
                try
                {
                    this.props.onFilterChange(this.state.items.length);
                }
                catch
                {

                }

            this.setState({
                filterChangeToken: this.props.filterChangeToken
            });
        }
    }

    private _processIncidents(uiStrings: Map<string, string>): void {
        var incidents: IPublicIncident[] = [];
        var groups: IGroup[] = [];
        var filteredItems: IPublicIncident[] = this.state.filter ? this.state.filter.filterItems(this._allItems) : this._allItems;

        if (this.props.onFilterChange)
            try {
                this.props.onFilterChange(filteredItems.length);
            }
            catch
            {

            }
        incidents = filteredItems;
        incidents = incidents.sort((a: IPublicIncident, b: IPublicIncident) => a.service > b.service ? 1 : -1);

        var count = 1;
        var currentGroup: IGroup = {
            name: "",
            key: "",
            startIndex: -1,
            count: -1,
            isCollapsed: false
        };
        for (var i = 0; i < incidents.length; i++) {
            if (currentGroup.name !== incidents[i].service) {
                if (currentGroup.startIndex !== -1) {
                    currentGroup.count = count;
                    groups.push(currentGroup);
                    count = 1;
                }

                currentGroup = {
                    name: incidents[i].service,
                    key: incidents[i].service,
                    startIndex: i,
                    count: 0,
                    isCollapsed: false
                }
            }
            else {
                count++
            }
        }

        if (currentGroup !== undefined) {
            currentGroup.count = count;
            groups.push(currentGroup);
        }

        this.setState({
            items: incidents,
            incidentGroups: groups,
            uiStrings: uiStrings,
            isDataLoaded: true
        });
    }

    _onOpenDetailsPanel(id: string) {
        this.setState({
            selectedIncident: id,
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
                    selectedIncident: this.state.items[itemIndex - 1].id
                });
            }
        }
    }

    private _selectNext(id: string): void {
        if (this.state.items !== undefined) {
            const itemIndex: number = this.state.items.findIndex((i) => i.id === id);

            if (itemIndex < this.state.items.length) {
                this.setState({
                    selectedIncident: this.state.items[itemIndex + 1].id
                });
            }
        }
    }

    private _isFirst(id: string): boolean {
        if (this.state.items !== undefined && this.state.items.length > 0) {
            const firstElement: IPublicIncident = this.state.items[0];
            return firstElement.id === id;
        } else
            return true;
    }

    private _isLast(id: string): boolean {
        if (this.state.items !== undefined && this.state.items.length > 0) {
            const lastElement: IPublicIncident = this.state.items[this.state.items.length - 1];
            return lastElement.id === id;
        } else
            return true;
    }
}