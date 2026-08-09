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
import { PublicMessageDetails } from './PublicMessageDetails';
import { ICustomAction, CustomAction } from '../../../CustomAction';
import { ServiceComponent } from '../../../ServiceNameComponent';
import { IPublicEvents, IPublicMessage } from '../../../IServiceStatistics';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { MessageBar, MessageBarType } from '@fluentui/react';
import { acquireAccessToken, accessControl } from "../../../../auth/AccessTokenHelper";
import { AuthenticationResult } from '@azure/msal-common';
import { checkInTeams } from '../../../auth/detectTeams';
import { AccessDenied } from "../../../AccessDenied";
import { translate, getUIString } from "../../../../translation/translator";
import { GlobalState } from "../../../GlobalState";

interface IPublicNewsState {
    isDetailsPanelOpen: boolean;
    selectedItem: string;
    inTeams: boolean;
    statistics?: IPublicEvents;
    items: IPublicMessage[];
    isDataLoaded: boolean;
    uiStrings: Map<string, string>;
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

export class PublicNews extends React.Component<{ collectionMode?: NewsItems }, IPublicNewsState> {
    _allNews: IPublicMessage[] = [];
    static contextType = GlobalState;

    constructor(props: { collectionMode?: NewsItems }) {
        super(props);

        var uiStrings: Map<string, string> = new Map([
            ["Title", "Title"],
            ["Couldn't retrieve data. Error:", "Couldn't retrieve data. Error:"],
            ["Filter", "Filter"],
            ["Close", "Close"],
            ["item", "item"],
            ["items", "items"],
            ["filtered", "filtered"],
            ["Updated", "Updated"],
            ["Clear filter", "Clear filter"],
            ["Select services", "Select services"],
            ["Service", "Service"],
            ["Services", "Services"],
            ["Sorted newer to older", "Sorted newer to older"],
            ["Sorted older to newer", "Sorted older to newer"],
            ["There are no published service updates.", "There are no published service updates."]
        ]);

        this.state = {
            items: [],
            isDetailsPanelOpen: false,
            selectedItem: "",
            isDataLoaded: false,
            inTeams: checkInTeams(),
            error: undefined,
            accessGranted: undefined,
            uiStrings: uiStrings,
            statistics: undefined
        };
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

        const { items, isDetailsPanelOpen, selectedItem, isDataLoaded, uiStrings, error } = this.state;

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
                onRender: (item: IPublicMessage) => {
                    return <div className="container" style={{ cursor: 'default' }} >
                            <div className="row" >
                                <div className="col" style={{ justifyContent: 'center', paddingLeft: '12px', paddingRight: '0px' }} >
                                    <FontIcon
                                    iconName='InfoSolid'
                                    className={classNames.advisory} />
                                </div>
                            </div>
                       </div>;
                },
                isPadded: true,
            },
            {
                key: 'clTitle',
                name: getUIString(uiStrings, "Title"),
                fieldName: 'title',
                minWidth: 180,
                isResizable: true,
                isCollapsible: false,
                isMultiline: false,
                data: 'string',
                onRender: (item: IPublicMessage) => {
                    return (<Link onClick={(event) => {
                                    event.preventDefault();
                                    this._onOpenDetailsPanel(item.id);
                    }} styles={pipeFabricStyles}>{item.title}</Link >);
                },
                isPadded: false,
            },
            {
                key: 'clServices',
                name: getUIString(uiStrings, "Service"),
                minWidth: 200,
                maxWidth: 200,
                isResizable: false,
                isCollapsible: true,
                isPadded: false,
                onRender: (item: IPublicMessage) => {
                    return <>
                        <TooltipHost content={item.services.join(", ")}><span>{
                            item.services.map((s) => (
                                <ServiceComponent name={s} />
                            ))
                        }</span></TooltipHost>
                        </>;
                },
            },
            {
                key: 'clUpdated',
                name: getUIString(uiStrings, "Updated"),
                minWidth: 60,
                maxWidth: 60,
                isResizable: false,
                isCollapsible: true,
                isPadded: false,
                onRender: (item: IPublicMessage) => {
                    return <>{item.lastModified ? new Date(item.lastModified).toLocaleDateString() : ""}</>;
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
                onRender: (item: IPublicMessage) => {          
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
                <PublicMessageDetails id={selectedItem} /> 
            </Panel>
        </>
        );
    }

    componentDidMount() {
        this._getNews();
    }

    private _getNews() {
        const requiredRoles: string[] = ['ServiceHealthReader', 'Communication.Write.All', 'Admin'];

        var translatedUIStrings: Map<string, string>;
        let globalState: any = this.context;
        var language: string = globalState.getProfileProperty("language", "en").trim();

        var langQuery: string = "";
        if (language !== "" && language !== "en") {
            langQuery = "?lang=" + language;
        }

        var authResponse: AuthenticationResult;

        acquireAccessToken()
            .then((response) => {
                authResponse = response;

                var tokenClaims: any = response.account?.idTokenClaims;
                const userRoles: any = tokenClaims?.roles;
                var userHasRequiredRole: boolean = userRoles.some((r: string) => requiredRoles.includes(r));

                this.setState({
                    accessGranted: userHasRequiredRole
                });

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
                            this._processMessages(result, translatedUIStrings);
                        }).catch((err) => {
                            this.setState({
                                error: err.message
                            });
                        });
                }
                else {
                    this._processMessages(this.state.statistics, translatedUIStrings);
                }
            });
    }

    private _processMessages(result: any, uiStrings: Map<string, string>): void {
        var statistics: IPublicEvents | undefined = undefined;
        var items: IPublicMessage[] = [];

        statistics = result;
        statistics!.messages = statistics!.messages.sort((a: any, b: any) => a.lastModified > b.lastModified ? -1 : 1);

        if (statistics !== undefined) {
            for (const message of statistics!.messages) {
                items.push({
                    id: message.id,
                    title: message.title,
                    services: message.services.length > 0 ? message.services : ["General announcement"],
                    lastModified: message.lastModified,
                    startTime: message.startTime,
                    comments: message.comments,
                    content: message.content
                });

                items = items.sort((a, b) => 0 - (a.lastModified > b.lastModified ? 1 : -1));
            }
        }

        this.setState({
            statistics: statistics,
            items: items,
            uiStrings: uiStrings,
            isDataLoaded: true
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
            const firstElement: IPublicMessage = this.state.items[0];
            return firstElement.id === id;
        } else
            return true;
    }

    private _isLast(id: string): boolean {
        if (this.state.items !== undefined && this.state.items.length > 0) {
            const lastElement: IPublicMessage = this.state.items[this.state.items.length - 1];
            return lastElement.id === id;
        } else
            return true;
    }
}