import * as React from 'react';
import { Text } from '@fluentui/react';
import { Image, Button, ButtonType, Separator } from '@fluentui/react';
import type { IContextualMenuItem, IContextualMenuProps } from '@fluentui/react';
import {
    DashboardCard,
    DashboardCardBodyText,
    ThumbnailItemStack,
    CompoundButtonStackV2
} from '@m365-admin/card';
import type { AddCardPanelInfoType, IAddCardPanelItemProps } from '@m365-admin/dashboard';
import { Dashboard } from '@m365-admin/dashboard';
import type {
    IRearrangeableGridCellDefinition,
    IRearrangeableGridDefinition,
    IRearrangeableGridGap,
} from '@m365-admin/rearrangeable-grid';
import {
    compactVerticalThenHorizontal,
    isFullGridGap,
} from '@m365-admin/rearrangeable-grid';
import type {
    IRecommendationV2StyleProps,
    IRecommendationV2Styles,
} from '@m365-admin/recommendation';
import { DashboardRecommendationWrapper } from '@m365-admin/recommendation';
import { EmptyStateImageSize } from '@m365-admin/empty-state';
import type { IActionBarProps } from '@m365-admin/action-bar';
import { ActionBar } from '@m365-admin/action-bar';
import { GlobalState } from '../../GlobalState';
import { HealthReport, HealthReportType, HealthReportMode } from './elements/Incidents';
import { News, NewsItems } from './elements/News';
import { translate, getUIString } from "../../../translation/translator";
import { IPublicEvents } from '../../IServiceStatistics';
import { acquireAccessToken } from "../../../auth/AccessTokenHelper";
import { AuthenticationResult } from '@azure/msal-browser';
import { AccessDenied } from '../../AccessDenied';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { PublicHealthReport } from './elements/PublicIncidents';
import { PublicNews } from './elements/PublicNews';

interface IDashboards {
    [key: string]: IRearrangeableGridDefinition | undefined
}

interface IPublicDashboardState {
    panelOpen: boolean;
    isLoading: boolean;
    useCompaction: boolean;
    uiStrings: Map<string, string>;
    map: IRearrangeableGridDefinition | undefined;
    dashboards: IDashboards;
    changeMessage: string;
    statistics?: IPublicEvents;
    error?: string;
    accessGranted?: boolean;
    isDataLoaded: boolean;
}

/* const resizeCallback =
    useRef<
        (
            cellWidth: number,
            cellHeight: number,
            itemKey: string,
            position?: { col: number; row: number },
        ) => void
    >();

const columnResizeRef =
    useRef<(width: number, height: number, gap: IRearrangeableGridGap) => void>();
const columnResize = (width: number, height: number, gap: IRearrangeableGridGap) => {
    columnResizeRef.current?.(width, height, gap);
}; */

enum CardSizes {
    small = 'Small',
    mediumTall = 'Medium (tall)',
    mediumWide = 'Medium (wide)',
    mediumUltrawide = 'Medium (ultra-wide)',
    large = 'Large',
}

type SizeMap = { [key in CardSizes]: { cellWidth: number; cellHeight: number } };

const CardSizeValues: SizeMap = {
    Small: { cellWidth: 1, cellHeight: 1 },
    'Medium (tall)': { cellWidth: 1, cellHeight: 2 },
    'Medium (wide)': { cellWidth: 2, cellHeight: 1 },
    'Medium (ultra-wide)': { cellWidth: 3, cellHeight: 1 },
    Large: { cellWidth: 2, cellHeight: 2 },
};

const removeText = 'Remove';
const resizeText = 'Resize';

const defaultLayout: IRearrangeableGridDefinition = {
    publicM365Issues: { cellHeight: 1, cellWidth: 3, row: 0, col: 0 },
    publishedM365Accouncements: { cellHeight: 1, cellWidth: 3, row: 1, col: 0 },
};

export class PublicDashboard extends React.Component<{}, IPublicDashboardState> {
    static contextType = GlobalState;
    private addCardPanelInfo: AddCardPanelInfoType = {};
;

    constructor(props: {}) {
        super(props);

        var uiStrings: Map<string, string> = new Map([
            ["Feature announcements", "Feature announcements"],
            ["Active incidents", "Active incidents"],
            ["Couldn't retrieve data. Error:", "Couldn't retrieve data. Error:"],
            ["Parts of this page may have been machine-translated.", "Parts of this page may have been machine-translated."],
            ["More actions", "More actions"],
            ["Public announcements", "Public announcements"],
            ["Use the arrow keys to navigate to a card.Press alt - shift - arrow to move a card.Press enter key to enter each item.Press escape to return to the grid when within an item.", "Use the arrow keys to navigate to a card.Press alt - shift - arrow to move a card.Press enter key to enter each item.Press escape to return to the grid when within an item."],
            ["Close add card panel", "Close add card panel"],
            ["Add cards to your home page", "Add cards to your home page"],
            ["Hold on, the data is on its way", "Hold on, the data is on its way"],
            ["Select a card, and then Enter to add the card to the bottom of the home page.", "Select a card, and then Enter to add the card to the bottom of the home page."],
            ["You already have all the cards on your home page", "You already have all the cards on your home page"],
            ["Published Microsoft 365 incidents and advisories", "Published Microsoft 365 incidents and advisories"],
            ["See a list of published Microsoft 365 incidents and advisories", "See a list of published Microsoft 365 incidents and advisories"],
            ["Card with list of published Microsoft 365 incidents and advisories", "Card with list of published Microsoft 365 incidents and advisories"],
            ["Add Published Microsoft 365 incidents and advisories card", "Add Published Microsoft 365 incidents and advisories card"],
            ["Published Microsoft 365 Message Center announcements", "Published Microsoft 365 Message Center announcements"],
            ["See a list of published Message Center announcements", "See a list of published Message Center announcements"],
            ["Card with list of published Message Center announcements", "Card with list of published Message Center announcements"],
            ["Add Published Microsoft 365 Message Center announcements card", "Add Published Microsoft 365 Message Center announcements card"],
            ["Remove", "Remove"],
            ["Remove card", "Remove card"],
            ["Resize", "Resize"]
        ]);

        this.state = {
            panelOpen: false,
            isLoading: false,
            useCompaction: true,
            uiStrings: uiStrings,
            map: defaultLayout,
            dashboards: {},
            changeMessage: "",
            isDataLoaded: false
        }

    }

    updateAddCardPanelInfo() {
        const { uiStrings } = this.state;

        this.addCardPanelInfo = {
            publicM365Issues: {
                itemProps: {
                    header: getUIString(uiStrings, "Published Microsoft 365 incidents and advisories"),
                    description:
                        getUIString(uiStrings, "See a list of published Microsoft 365 incidents and advisories"),
                    imageSrc:
                        '/images/addCard/detailsList.svg',
                    imageAlt: getUIString(uiStrings, "Card with list of published Microsoft 365 incidents and advisories"),
                    tooltipHostProps: { content: getUIString(uiStrings, "Add Published Microsoft 365 incidents and advisories card"), },
                },
                height: 1,
                width: 3,
            },
            publishedM365Accouncements: {
                itemProps: {
                    header: getUIString(uiStrings, "Published Microsoft 365 Message Center announcements"),
                    description:
                        getUIString(uiStrings, "See a list of published Message Center announcements"),
                    imageSrc:
                        '/images/addCard/detailsList.svg',
                    imageAlt: getUIString(uiStrings, "Card with list of published Message Center announcements"),
                    tooltipHostProps: {
                        content: getUIString(uiStrings, "Add Published Microsoft 365 Message Center announcements card"),
                    }
                },
                height: 1,
                width: 3,
            }
        };
    }

    public render() {
        const { panelOpen, isLoading, useCompaction, map, uiStrings, statistics, accessGranted, isDataLoaded } = this.state;
        var translatedString: string = getUIString(uiStrings, "Parts of this page may have been machine-translated.");
        const translated: boolean = translatedString !== "Parts of this page may have been machine-translated.";
        console.log(uiStrings);
        if (accessGranted !== undefined && accessGranted === false)
            return (<AccessDenied />);

        if (!isDataLoaded)
            return (<div className="loadingProgress">
                <br />
                <Spinner size={SpinnerSize.large} />
                <br />&nbsp;
            </div>);

        return (
            <div className="container" style={{ paddingTop: '15px', paddingBottom: '15px' }}>
                <div className="row">
                    <div className="col">
                        <Text variant="mediumPlus"><b>{getUIString(uiStrings, "Public announcements")}</b></Text>
                        <Separator />
                    </div>
                </div>
                <div className="row">
                    <div className="col">
                        <Button buttonType={ButtonType.command} iconProps={{ iconName: 'Add' }} onClick={() => this._openPanel()}>Add cards</Button>
                    </div>
                </div>

                <div className="row" style={{marginTop: '24px'}} >
                        <div className="col">
                        <Dashboard
                            isAddCardPanelOpen={panelOpen}
                            map={map}
                            onChange={this._onChange}
                            onAddRemove={this._onAddRemove}
                            ariaLabel={getUIString(uiStrings, "Public announcements")}
                            ariaDescription={getUIString(uiStrings, "Use the arrow keys to navigate to a card.Press alt - shift - arrow to move a card.Press enter key to enter each item.Press escape to return to the grid when within an item.")}
                            rearrangeableGridProps={{

                                ...(useCompaction && { onCompaction: compactVerticalThenHorizontal }),
                            }}
                            panelProps={{
                                closeButtonAriaLabel: getUIString(uiStrings, "Close add card panel"),
                                headerText: getUIString(uiStrings, "Add cards to your home page"),
                            loadingProps: {
                                isLoading: isLoading,
                                primaryLoadingText: getUIString(uiStrings, "Hold on, the data is on its way"),
                            },
                            onDismiss: this._closePanel,
                            'aria-label':
                                getUIString(uiStrings, "Select a card, and then Enter to add the card to the bottom of the home page."),
                        }}
                        addCardPanelInfo={this.addCardPanelInfo}
                        emptyStateProps={{
                            body: getUIString(uiStrings, "You already have all the cards on your home page"),
                            nativeImgProps: {
                                /**
                                 * demo image, please work with your designers to pull in the correct assets and do not link to this store
                                 * for assets. Use a CDN!1
                                 */
                                src: '/images/addCard/verticalBarEmptyState.png',
                                ['aria-hidden']: 'true',
                            },
                            illustrationSize: EmptyStateImageSize.Large,
                        }}
                    >

                        <DashboardCard
                            key="publicM365Issues"
                            titleText={getUIString(uiStrings, "Active incidents")}
                            moreIconButtonProps={{
                                ariaLabel: getUIString(uiStrings, 'More actions'),
                                menuProps: this._generateItemMenu('publicM365Issues'),
                            }}
                            moreIconButtonTooltipHostProps={{ content: getUIString(uiStrings, 'More actions') }}
                            styles={{
                                secondaryBody: {
                                    // This effectively makes the right size a minimum width of 282,
                                    // before reflowing to a vertical orientation
                                    flex: '1 1 282px',
                                },
                            }}
                                body={
                                <>
                                        <PublicHealthReport type={HealthReportType.active} mode={HealthReportMode.dashboard} statistics={statistics} />
                                </>
                            }
                        />

                            <DashboardCard
                                key="publishedM365Accouncements"
                                titleText={getUIString(uiStrings, "Feature announcements")}
                                moreIconButtonProps={{
                                    ariaLabel: getUIString(uiStrings, 'More actions'),
                                    menuProps: this._generateItemMenu('publishedM365Accouncements'),
                                }}
                                moreIconButtonTooltipHostProps={{ content: getUIString(uiStrings, 'More actions') }}
                                styles={{
                                    secondaryBody: {
                                        // This effectively makes the right size a minimum width of 282,
                                        // before reflowing to a vertical orientation
                                        flex: '1 1 282px',
                                    },
                                }}
                                body={
                                    <>
                                        <PublicNews />
                                    </>
                                }
                            />
                    </Dashboard>
                        </div>
                    </div>
                </div>
        );
    }

    componentDidMount() {
        let globalState: any = this.context;
        var dashboardMap: IRearrangeableGridDefinition | undefined = globalState.getDashboard("M365Public");

        const requiredRoles: string[] = ['ServiceHealthReader', 'Communication.Write.All', 'Public', 'Admin', 'LicenseReader'];

        var translatedUIStrings: Map<string, string>;
        var statistics: IPublicEvents | undefined = undefined;
        var language: string = globalState.getProfileProperty("language", "en").trim();
        var langQuery: string = "";
        if (language !== "" && language !== "en") {
            langQuery = "?lang=" + language;
        }

        var authResponse: AuthenticationResult;
        var userHasRequiredRole: boolean = false;

        acquireAccessToken()
            .then((response) => {
                var tokenClaims: any = response.account?.idTokenClaims;
                const userRoles: any = tokenClaims?.roles;
                userHasRequiredRole = userRoles.some((r: string) => requiredRoles.includes(r));

                this.setState({
                    accessGranted: userHasRequiredRole
                });

                authResponse = response;

                if (userHasRequiredRole)
                    translate(language,
                        this.state.uiStrings,
                        authResponse
                    ).then((response) => {
                        translatedUIStrings = response;
                    });
            }).then(() => {
                if (userHasRequiredRole)
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
                            statistics = result;
                            this.setState({
                                statistics: statistics,
                                uiStrings: translatedUIStrings,
                                map: dashboardMap ? dashboardMap : defaultLayout,
                                isDataLoaded: true
                            });

                            this.updateAddCardPanelInfo();
                        }).catch((err) => {
                            this.setState({
                                map: dashboardMap ? dashboardMap : defaultLayout,
                                error: err.message
                            });
                        });
            });
    }

    private _openPanel = (): void => {
        this.setState({
            panelOpen: true
        });
    }

    private _closePanel = (): void => {
        this.setState({
            panelOpen: false
        });
    }

    private _onChange = (
        newMap: IRearrangeableGridDefinition | undefined,
        diffMap: IRearrangeableGridDefinition | undefined,
    ): void => {
        this.setState({
            map: newMap
        });

        let globalState: any = this.context;
        globalState.setDashboard("M365Public", newMap);
        
        if (diffMap) {
            const keys = Object.keys(diffMap);
            const changeString = keys.map((key: string) => {
                const cell: IRearrangeableGridCellDefinition = diffMap[key];
                const panelInfo = this.addCardPanelInfo[key]?.itemProps.header;

                if (!panelInfo) {
                    return null;
                }

                return (
                    <>{`${panelInfo} card has moved to row ${cell.row ?? 0 + 1} and column ${cell.col ?? 0 + 1
                        }`}</>
                );
            });

            this.setState({
                changeMessage: changeString.toString()
            });
        }
    }

    private _generateRemoveCallback = (key: string) => {
        return () => {
            const newMap = this.state.map;

            if (newMap !== undefined) {
                delete newMap[key];

                let globalState: any = this.context;
                globalState.setDashboard("M365Public",newMap);

                this.setState({
                    map: newMap
                });
            }
        };
    };

    private _generateResizeCallback = (key: string) => {
        return (
            ev: React.MouseEvent<HTMLElement, MouseEvent> | React.KeyboardEvent<HTMLElement> | undefined,
            item: IContextualMenuItem | undefined,
        ) => {
            if (item !== undefined) {
                const dimension = CardSizeValues[item.key as CardSizes];
                /* resizeCallback.current?.(dimension.cellWidth, dimension.cellHeight, key);
                let globalState: any = this.context;
                globalState.setDashboard("M365Public", newMap); */
            }
        };
    };

    private _generateItemMenu = (key: string): IContextualMenuProps => {
        return {
            items: [
                {
                    iconProps: { iconName: 'Cancel' },
                    key: "remove-" + key,
                    text: getUIString(this.state.uiStrings, "Remove"),
                    onClick: this._generateRemoveCallback(key),
                    ariaLabel: getUIString(this.state.uiStrings, "Remove card"),
                } /*,
                {
                    iconProps: { iconName: 'MiniExpand' },
                    key: resizeText,
                    text: resizeText,
                    subMenuProps: {
                        items: [
                            { key: CardSizes.small, text: CardSizes.small },
                            { key: CardSizes.mediumTall, text: CardSizes.mediumTall },
                            { key: CardSizes.mediumWide, text: CardSizes.mediumWide },
                            { key: CardSizes.large, text: CardSizes.large },
                        ],

                        onItemClick: this._generateResizeCallback(key),
                    },
                }, */
            ],
        };
    };

    private _onAddRemove = (
        additions: IAddCardPanelItemProps[],
        removals: IAddCardPanelItemProps[],
    ) => {

        if (additions.length > 0) {
            let globalState: any = this.context;
            globalState.setDashboard("M365Public", this.state.map);
        }

        const additionString =
            additions.length > 0
                ? `${additions
                    .map((item: IAddCardPanelItemProps) => {
                        return item.header;
                    })
                    .toString()} ${additions.length > 1 ? 'cards have' : 'card has'} been added`
                : '';

        const removalString =
            removals.length > 0
                ? `${removals
                    .map((item: IAddCardPanelItemProps) => {
                        return item.header;
                    })
                    .toString()} ${removals.length > 1 ? 'cards have' : 'card has'} been removed`
                : '';

        this.setState({
            changeMessage: additionString + removalString
        });
    };
}