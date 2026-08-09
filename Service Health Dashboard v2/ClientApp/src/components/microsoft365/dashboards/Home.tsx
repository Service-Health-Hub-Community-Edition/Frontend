import * as React from 'react';
import { Text } from '@fluentui/react';
import { ThemeProvider, ITheme } from '@fluentui/react';
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
import { D365PPReport, D365PPReportItems, D365PPReportMode } from './elements/D365PPReleases';
import { ActiveIncidentsGraph } from './elements/ActiveIncidentsGraph';
import { WeeklyGraph, WeeklyReportType } from './elements/WeeklyGraph';
import { ServiceStatisticsGraph } from './elements/ServiceStatisticsGraph';
import { IncidentStatisticsGraph } from './elements/IncidentStatisticsGraph';
import { MonthlyActiveUsersGraph } from './elements/MonthlyActiveUsersGraph';

interface IDashboards {
    [key: string]: IRearrangeableGridDefinition | undefined
}

interface IHomeDashboardState {
    panelOpen: boolean;
    isLoading: boolean;
    useCompaction: boolean;
    map: IRearrangeableGridDefinition | undefined;
    dashboards: IDashboards;
    changeMessage: string;
    theme: any;
    initialized: boolean;
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

const addCardPanelInfo: AddCardPanelInfoType = {
    activeIncidentsGraph: {
        itemProps: {
            header: 'Active incidents per service',
            description: "Shows number of active incidents per service",
            imageSrc:
                '/images/addCard/rbac.svg',
            imageAlt: 'Card with text',
            tooltipHostProps: { content: 'Add Active incidents per service card' },
        },
        height: 1,
        width: 1,
    },
    activeM365Issues: {
        itemProps: {
            header: 'Active Microsoft 365 issues',
            description:
                'See a list of active Microsoft 365 issues',
            imageSrc:
                '/images/addCard/detailsList.svg',
            imageAlt: 'Card with list of Microsoft 365 issues',
            tooltipHostProps: { content: 'Add Active Microsoft 365 issues card' },
        },
        height: 1,
        width: 2,
    },
    is180: {
        itemProps: {
            header: 'Incident statistics',
            description:
                "See distribution of incidents over category for the past 180 days",
            imageSrc:
                '/images/addCard/graph.svg',
            imageAlt: 'Incident statistics for past 180 days',
            tooltipHostProps: { content: 'Add Incident statistics card' },
        },
        height: 1,
        width: 1,
    },
    monthlyActiveUsers: {
        itemProps: {
            header: 'Microsoft 365 active users report',
            description: 'See how many users with Office 365 apps installed have used at least one of those apps recently',
            imageSrc:
                '/images/addCard/graph.svg',
            imageAlt: 'Microsoft 365 active users report',
            tooltipHostProps: { content: 'Add Microsoft 365 active users report card' },
        },
        height: 1,
        width: 2,
    },
    messageCenterWeekly: {
        itemProps: {
            header: 'Message Center weekly report',
            description: 'Showing distribution of message center communication over past 12 weeks',
            imageSrc:
                '/images/addCard/graph.svg',
            imageAlt: 'Message Center weekly report',
            tooltipHostProps: { content: 'Add Message Center weekly report card' },
        },
        height: 1,
        width: 2,
    },
    m365News: {
        itemProps: {
            header: 'Last week in Message Center',
            description:
                'See a list of new Message Center communications',
            imageSrc:
                '/images/addCard/detailsList.svg',
            imageAlt: 'Card with list of new Message Center communications',
            tooltipHostProps: { content: 'Add Last week in Message Center card' },
        },
        height: 1,
        width: 2,
    },
    d365Releases: {
        itemProps: {
            header: 'Dynamics 365 and Power Platform releases',
            description:
                'See a list of new Dynamics 365 and Power Platform releases',
            imageSrc:
                '/images/addCard/detailsList.svg',
            imageAlt: 'Card with list of new Dynamics 365 and Power Platform releases',
            tooltipHostProps: { content: 'Add Dynamics 365 and Power Platform releases card' },
        },
        height: 1,
        width: 2,
    },
    messageCenterPastWeek: {
        itemProps: {
            header: 'Message Center weekly report',
            description: 'Showing amount of Message Center communications per service, for past 7 days',
            imageSrc:
                '/images/addCard/rbac.svg',
            imageAlt: 'Message Center weekly report',
            tooltipHostProps: { content: 'Add Message Center weekly report card' },
        },
        height: 1,
        width: 1,
    },
    trainingsAndResources: {
        itemProps: {
            header: 'Trainings and ressources',
            description: 'Get guidance on managing Microsoft 365 changes and incidents',
            imageSrc:
                '/images/addCard/thumbnailItemStack.svg',
            imageAlt: 'Trainings and ressources',
            tooltipHostProps: { content: 'Add Trainings and ressources card' },
        },
        height: 1,
        width: 1,
    }
};

export class HomeDashboard extends React.Component<{}, IHomeDashboardState> {
    static contextType = GlobalState;

    constructor(props: {}) {
        super(props);

        this.state = {
            panelOpen: false,
            isLoading: false,
            useCompaction: true,
            map: {
                activeIncidentsGraph: { cellHeight: 1, cellWidth: 1, row: 0, col: 0 },
                activeM365Issues: { cellHeight: 1, cellWidth: 2, row: 0, col:1 },              
                is180: { cellHeight: 1, cellWidth: 1, row: 1, col: 0 },
                monthlyActiveUsers: { cellHeight: 1, cellWidth: 2, row: 1, col: 1 },
                messageCenterWeekly: { cellHeight: 1, cellWidth: 2, row: 2, col: 0 },
                d365Releases: { cellHeight: 1, cellWidth: 2, row: 3, col: 0 },
                messageCenterPastWeek: { cellHeight: 1, cellWidth: 1, row: 2, col: 1 },
            },
            dashboards: {},
            changeMessage: "",
            theme: undefined,
            initialized: false
        }
    }

    public render() {
        const { panelOpen, isLoading, useCompaction, map, theme, initialized } = this.state;

        if (!initialized) {
            return (<></>)
        }

        return (
            <ThemeProvider theme={this.state.theme}>
            <div className="container" style={{ paddingTop: '15px', paddingBottom: '15px' }}>
                <div className="row">
                    <div className="col">
                        <Text variant="mediumPlus"><b>Microsoft 365</b></Text>
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
                        ariaLabel="Microsoft 365"
                        ariaDescription="Use the arrow keys to navigate to a card. Press alt-shift-arrow to move a card. Press enter key to enter each item. Press escape to return to the grid when within an item."
                        rearrangeableGridProps={{

                            ...(useCompaction && { onCompaction: compactVerticalThenHorizontal }),
                        }}
                        panelProps={{
                            closeButtonAriaLabel: 'Close add card panel',
                            headerText: 'Add cards to your home page',
                            loadingProps: {
                                isLoading: isLoading,
                                primaryLoadingText: 'Hold on, the data is on its way',
                            },
                            onDismiss: this._closePanel,
                            'aria-label':
                                'Select a card, and then Enter to add the card to the bottom of the home page.',
                        }}
                        addCardPanelInfo={addCardPanelInfo}
                        emptyStateProps={{
                            body: 'You already have all the cards on your home page',
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
                            key="activeM365Issues"
                            titleText="Active issues"
                            moreIconButtonProps={{
                                ariaLabel: 'More actions',
                                menuProps: this._generateItemMenu('activeM365Issues'),
                            }}
                            moreIconButtonTooltipHostProps={{ content: 'More actions' }}
                            styles={{
                                secondaryBody: {
                                    // This effectively makes the right size a minimum width of 282,
                                    // before reflowing to a vertical orientation
                                    flex: '1 1 282px',
                                },
                            }}
                                body={
                                <>
                                    <HealthReport type={HealthReportType.active} mode={HealthReportMode.dashboard} />
                                </>
                            }

                            footer={
                                <ActionBar primaryButtonProps={{
                                    text: 'View more',
                                    href: '/microsoft365/health'
                                }} />
                            }
                        />

                        <DashboardCard
                            key="activeIncidentsGraph"
                            headerText={<><br/>Active issues</>}
                            titleText="Service issues"
                            moreIconButtonProps={{
                                ariaLabel: 'More actions',
                                menuProps: this._generateItemMenu('activeIncidentsGraph'),
                            }}
                            moreIconButtonTooltipHostProps={{ content: 'More actions' }}
                            body={
                                <div>
                                    <div className="container" style={{ padding: '0px' }}>
                                        <div className="row">
                                            <div className="col ml-auto" style={{ marginLeft: '0px', marginRight: '0px' }} >
                                                <ActiveIncidentsGraph />
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            }
                            />

                            <DashboardCard
                                key="monthlyActiveUsers"
                                titleText="Microsoft 365 active users report"
                                moreIconButtonProps={{
                                    ariaLabel: 'More actions',
                                    menuProps: this._generateItemMenu('monthlyActiveUsers'),
                                }}
                                moreIconButtonTooltipHostProps={{ content: 'More actions' }}
                                body={
                                    <div style={{ display: 'flex', width: '100%' }} >
                                        <div className="container"
                                            id="mauContainer"
                                            style={{
                                                padding: '0px',
                                                overflow: 'hidden',
                                                flex: '1 1 auto',
                                                width: document.querySelector('[data-key="monthlyActiveUsers"]')?.getBoundingClientRect().width,
                                                position: 'relative',
                                                display: 'block'
                                            }}
                                        >
                                            <div className="row" style={{
                                                display: 'flex',
                                                minHeight: '100%',
                                                flexDirection: 'column'
                                            }} >
                                                <div className="col" style={{
                                                    marginLeft: '0px', marginRight: '0px',
                                                    display: 'flex',
                                                    flexFlow: 'row wrap',
                                                    flex: '1 1 auto',
                                                    overflow: 'hidden',
                                                    wordBreak: 'break-word'
                                                }} >
                                                    <MonthlyActiveUsersGraph container="mauContainer" />
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                }
                            />

                            <DashboardCard
                                key="is180"
                                titleText="Incidents and advisories breakdown"
                                moreIconButtonProps={{
                                    ariaLabel: 'More actions',
                                    menuProps: this._generateItemMenu('is180'),
                                }}
                                moreIconButtonTooltipHostProps={{ content: 'More actions' }}
                                body={
                                    <div style={{ display: 'flex', width: '100%' }} >
                                        <div className="container"
                                            id="incidentStatistics180Container"
                                            style={{
                                                padding: '0px',
                                                overflow: 'hidden',
                                                flex: '1 1 auto',
                                                width: document.querySelector('[data-key="is180"]')?.getBoundingClientRect().width,
                                                position: 'relative',
                                                display: 'block'
                                            }}
                                        >
                                            <div className="row" style={{
                                                display: 'flex',
                                                minHeight: '100%',
                                                flexDirection: 'column'
                                            }} >
                                                <div className="col" style={{
                                                    marginLeft: '0px', marginRight: '0px',
                                                    display: 'flex',
                                                    flexFlow: 'row wrap',
                                                    flex: '1 1 auto',
                                                    overflow: 'hidden',
                                                    wordBreak: 'break-word'
                                                }} >
                                                    <IncidentStatisticsGraph container="incidentStatistics180Container" />
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                }
                                footer={
                                    <ActionBar primaryButtonProps={{
                                        text: 'View more',
                                        href: '/reports'
                                    }} />
                                }
                            />

                            <DashboardCard
                                key="messageCenterWeekly"
                                headerText={<><br />Message Center items per week</>}
                                titleText="Message Center weekly report"
                                moreIconButtonProps={{
                                    ariaLabel: 'More actions',
                                    menuProps: this._generateItemMenu('messageCenterWeekly'),
                                }}
                                moreIconButtonTooltipHostProps={{ content: 'More actions' }}
                                body={
                                    <div style={{ display: 'flex', width: '100%' }} >
                                        <div className="container"
                                            id="mcWeeklyContainer"
                                            style={{
                                                padding: '0px',
                                                overflow: 'hidden',
                                                flex: '1 1 auto',
                                                width: document.querySelector('[data-key="messageCenterWeekly"]')?.getBoundingClientRect().width,
                                                position: 'relative',
                                                display: 'block'
                                            }}
                                        >
                                        <div className="row" style={{
                                            display: 'flex',
                                            minHeight: '100%',
                                            flexDirection: 'column'
                                        }} >
                                            <div className="col" style={{
                                                marginLeft: '0px', marginRight: '0px',
                                                display: 'flex',
                                                flexFlow: 'row wrap',
                                                flex: '1 1 auto',
                                                overflow: 'hidden',
                                                wordBreak: 'break-word'
                                            }} >
                                                    <WeeklyGraph type={WeeklyReportType.MessageCenter} container="mcWeeklyContainer"/>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                }
                            />

                            <DashboardCard
                                key="m365News"
                                titleText="Last week in Message Center"
                                moreIconButtonProps={{
                                    ariaLabel: 'More actions',
                                    menuProps: this._generateItemMenu('m365News'),
                                }}
                                moreIconButtonTooltipHostProps={{ content: 'More actions' }}
                                styles={{
                                    secondaryBody: {
                                        // This effectively makes the right size a minimum width of 282,
                                        // before reflowing to a vertical orientation
                                        flex: '1 1 282px',
                                    },
                                }}
                                body={
                                    <>
                                        <News collectionMode={NewsItems.Inbox} />
                                    </>
                                }

                                footer={
                                    <ActionBar primaryButtonProps={{
                                        text: 'View more',
                                        href: '/messages'
                                    }} />
                                }
                            />

                            <DashboardCard
                                key="d365Releases"
                                titleText="Dynamics 365 and Power Platform releases"
                                moreIconButtonProps={{
                                    ariaLabel: 'More actions',
                                    menuProps: this._generateItemMenu('d365Releases'),
                                }}
                                moreIconButtonTooltipHostProps={{ content: 'More actions' }}
                                styles={{
                                    secondaryBody: {
                                        // This effectively makes the right size a minimum width of 282,
                                        // before reflowing to a vertical orientation
                                        flex: '1 1 282px',
                                    },
                                }}
                                body={
                                    <>
                                        <D365PPReport collectionMode={D365PPReportItems.Inbox} mode={D365PPReportMode.dashboard} />
                                    </>
                                }

                                footer={
                                    <ActionBar primaryButtonProps={{
                                        text: 'View more',
                                        href: '/d365pp'
                                    }} />
                                }
                            />

                            <DashboardCard
                                key="messageCenterPastWeek"
                                headerText={<><br />Message Center items</>}
                                titleText="Message Center (past 7 days)"
                                moreIconButtonProps={{
                                    ariaLabel: 'More actions',
                                    menuProps: this._generateItemMenu('messageCenterPastWeek'),
                                }}
                                moreIconButtonTooltipHostProps={{ content: 'More actions' }}
                                body={
                                    <div className="container"
                                        id="mcWeeklyContainer"
                                        style={{
                                            padding: '0px',
                                            overflow: 'hidden'
                                        }}>
                                        <div className="row">
                                            <div className="col" style={{ marginLeft: '0px', marginRight: '0px' }} >
                                                <ServiceStatisticsGraph />
                                            </div>
                                        </div>
                                    </div>
                                }
                            />

                            <DashboardCard
                                key="trainingsAndResources"
                                titleText="Trainigs and resources"
                                moreIconButtonProps={{
                                    ariaLabel: 'More actions',
                                    menuProps: this._generateItemMenu('trainingsAndResources'),
                                }}
                                moreIconButtonTooltipHostProps={{ content: 'More actions' }}
                                body={
                                    <ThumbnailItemStack
                                        items={[
                                            {
                                                imageProps: {
                                                    src: '/images/addCard/domains.svg',
                                                    alt: 'Microsoft 365 service health',
                                                },
                                                compoundButtonProps: {
                                                    text: 'Microsoft 365 service health',
                                                    secondaryText: 'Microsoft 365 service health documentation',
                                                    href: 'https://learn.microsoft.com/en-us/microsoft-365/enterprise/view-service-health?view=o365-worldwide',
                                                    target: '_blank'
                                                },
                                            },
                                            {
                                                imageProps: {
                                                    src: '/images/adminTraining.svg',
                                                    alt: 'Microsoft 365 change guide',
                                                },
                                                compoundButtonProps: {
                                                    text: 'Microsoft 365 change guide',
                                                    secondaryText: 'Best practices from the field',
                                                    href: 'https://learn.microsoft.com/en-us/deployoffice/fieldnotes/microsoft-365-change-guide',
                                                    target: '_blank'
                                                },
                                            },
                                            {
                                                imageProps: {
                                                    src: '/images/userTraining.svg',
                                                    alt: 'Microsoft 365 Blog',
                                                },
                                                compoundButtonProps: {
                                                    primary: false,
                                                    text: 'Microsoft 365 Blog',
                                                    secondaryText: 'Recent news and articles around Microsoft 365',
                                                    href: 'https://www.microsoft.com/en-us/microsoft-365/blog/',
                                                    target: '_blank'
                                                },
                                            },
                                            {
                                                imageProps: {
                                                    src: '/images/customize.svg',
                                                    alt: 'Options, tools and services with their respective audience',
                                                },
                                                compoundButtonProps: {
                                                    primary: false,
                                                    text: 'Office 365 Update Scout',
                                                    secondaryText: 'Options, tools and services with their respective audience',
                                                    href: 'https://aka.ms/o365UpdateScout',
                                                    target: '_blank'
                                                },
                                            },
                                        ]} />
                                }
                            />
                    </Dashboard>
                        </div>
                    </div>
                </div>
            </ThemeProvider>
        );
    }

    componentDidMount() {
        let globalState: any = this.context;
        var dashboardMap: IRearrangeableGridDefinition | undefined = globalState.getDashboard("M365Home");

        var theme: ITheme = globalState.getTheme();

        this.setState({
            map: dashboardMap ? dashboardMap : {
                activeIncidentsGraph: { cellHeight: 1, cellWidth: 1, row: 0, col: 0 },
                activeM365Issues: { cellHeight: 1, cellWidth: 2, row: 0, col: 1 },
                is180: { cellHeight: 1, cellWidth: 1, row: 1, col: 0 },
                monthlyActiveUsers: { cellHeight: 1, cellWidth: 2, row: 1, col: 1 },
                messageCenterWeekly: { cellHeight: 1, cellWidth: 2, row: 2, col: 0 },
                messageCenterPastWeek: { cellHeight: 1, cellWidth: 1, row: 2, col: 1 },
            },
            initialized: true,
            theme: theme
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
        globalState.setDashboard("M365Home", newMap);
        
        if (diffMap) {
            const keys = Object.keys(diffMap);
            const changeString = keys.map((key: string) => {
                const cell: IRearrangeableGridCellDefinition = diffMap[key];
                const panelInfo = addCardPanelInfo[key]?.itemProps.header;

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
                globalState.setDashboard("M365Home",newMap);

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
                globalState.setDashboard("M365Home", newMap); */
            }
        };
    };

    private _generateItemMenu = (key: string): IContextualMenuProps => {
        return {
            items: [
                {
                    iconProps: { iconName: 'Cancel' },
                    key: removeText,
                    text: removeText,
                    onClick: this._generateRemoveCallback(key),
                    ariaLabel: 'Remove card',
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
            globalState.setDashboard("M365Home", this.state.map);
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