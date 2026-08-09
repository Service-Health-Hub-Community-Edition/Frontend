
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
import { AzureUpdatesReport, AzureUpdatesReportType, AzureUpdatesReportMode } from './elements/Updates';
import { ActiveIncidentsGraph } from './elements/ActiveIncidentsGraph';

interface IDashboards {
    [key: string]: IRearrangeableGridDefinition | undefined
}

interface IAzureHomeDashboardState {
    panelOpen: boolean;
    isLoading: boolean;
    useCompaction: boolean;
    map: IRearrangeableGridDefinition | undefined;
    dashboards: IDashboards;
    changeMessage: string;
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
    azureActiveAlertsGraph: {
        itemProps: {
            header: 'Active alerts per service',
            description: "Shows number of active alerts per service",
            imageSrc:
                '/images/addCard/rbac.svg',
            imageAlt: 'Card with text',
            tooltipHostProps: { content: 'Add Active alerts per service card' },
        },
        height: 1,
        width: 1,
    },
    azureActiveAlertsList: {
        itemProps: {
            header: 'Active Azure alerts',
            description:
                'See a list of active Azure alerts',
            imageSrc:
                '/images/addCard/detailsList.svg',
            imageAlt: 'Card with list of Azure alerts',
            tooltipHostProps: { content: 'Add Active Azure alerts card' },
        },
        height: 1,
        width: 2,
    },
    azureActiveUpdatesList: {
        itemProps: {
            header: 'Upcoming Azure updates',
            description:
                'See a list of upcoming Azure updates',
            imageSrc:
                '/images/addCard/detailsList.svg',
            imageAlt: 'Card with list of Azure updates',
            tooltipHostProps: { content: 'Add Active Azure updates card' },
        },
        height: 1,
        width: 2,
    },
    azureTrainingsAndResources: {
        itemProps: {
            header: 'Trainings and ressources',
            description: 'Get guidance on managing Azure incidents',
            imageSrc:
                '/images/addCard/thumbnailItemStack.svg',
            imageAlt: 'Trainings and ressources',
            tooltipHostProps: { content: 'Add Trainings and ressources card' },
        },
        height: 1,
        width: 1,
    }
};

export class AzureHomeDashboard extends React.Component<{}, IAzureHomeDashboardState> {
    static contextType = GlobalState;

    constructor(props: {}) {
        super(props);

        this.state = {
            panelOpen: false,
            isLoading: false,
            useCompaction: true,
            map: {
                azureActiveAlertsGraph: { cellHeight: 1, cellWidth: 1, row: 0, col: 0 },
                azureActiveAlertsList: { cellHeight: 1, cellWidth: 2, row: 0, col: 1 }
            },
            dashboards: {},
            changeMessage: ""
        }

    }

    public render() {
        const { panelOpen, isLoading, useCompaction, map } = this.state;

        return (
            <div className="container" style={{ paddingTop: '15px', paddingBottom: '15px' }}>
                <div className="row">
                    <div className="col">
                        <Text variant="mediumPlus"><b>Microsoft Azure</b></Text>
                        <Separator />
                    </div>
                </div>
                <div className="row">
                    <div className="col">
                        <Button buttonType={ButtonType.command} iconProps={{ iconName: 'Add' }} onClick={() => this._openPanel()}>Add cards</Button>
                    </div>
                </div>

                <div className="row" style={{ marginTop: '24px' }} >
                    <div className="col">
                        <Dashboard
                            isAddCardPanelOpen={panelOpen}
                            map={map}
                            onChange={this._onChange}
                            onAddRemove={this._onAddRemove}
                            ariaLabel="Microsoft Azure"
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
                                key="azureActiveAlertsList"
                                titleText="Active alerts"
                                moreIconButtonProps={{
                                    ariaLabel: 'More actions',
                                    menuProps: this._generateItemMenu('azureActiveAlertsList'),
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
                                        href: '/azure/health'
                                    }} />
                                } />

                            <DashboardCard
                                key="azureActiveUpdatesList"
                                titleText="Upcoming Azure updates"
                                moreIconButtonProps={{
                                    ariaLabel: 'More actions',
                                    menuProps: this._generateItemMenu('azureActiveUpdatesList'),
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
                                        <AzureUpdatesReport type={AzureUpdatesReportType.active} mode={AzureUpdatesReportMode.dashboard} />
                                    </>
                                }
                                footer={
                                    <ActionBar primaryButtonProps={{
                                        text: 'View more',
                                        href: '/azure/updates'
                                    }} />
                                } />

                            <DashboardCard
                                key="azureActiveAlertsGraph"
                                headerText={<><br />Active alerts</>}
                                titleText="Azure alerts"
                                moreIconButtonProps={{
                                    ariaLabel: 'More actions',
                                    menuProps: this._generateItemMenu('azureActiveAlertsGraph'),
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
                                key="azureTrainingsAndResources"
                                titleText="Trainigs and resources"
                                moreIconButtonProps={{
                                    ariaLabel: 'More actions',
                                    menuProps: this._generateItemMenu('azureTrainingsAndResources'),
                                }}
                                moreIconButtonTooltipHostProps={{ content: 'More actions' }}
                                body={
                                    <ThumbnailItemStack
                                        items={[
                                            {
                                                imageProps: {
                                                    src: '/images/adminTraining.svg',
                                                    alt: 'Azure Service Health alerts',
                                                },
                                                compoundButtonProps: {
                                                    text: 'Azure Service Health alerts',
                                                    secondaryText: 'Configure Azure Monitor to receive alerts',
                                                    href: 'https://learn.microsoft.com/en-us/azure/service-health/alerts-activity-log-service-notifications-portal',
                                                    target: '_blank'
                                                },
                                            },
                                            {
                                                imageProps: {
                                                    src: '/images/customize.svg',
                                                    alt: 'Azure status',
                                                },
                                                compoundButtonProps: {
                                                    primary: false,
                                                    text: 'Azure status',
                                                    secondaryText: 'Public Azure status page',
                                                    href: 'https://azure.status.microsoft/en-us/status',
                                                    target: '_blank'
                                                },
                                            },
                                            {
                                                imageProps: {
                                                    src: '/images/userTraining.svg',
                                                    alt: 'Azure Updates',
                                                },
                                                compoundButtonProps: {
                                                    primary: false,
                                                    text: 'Azure Updates',
                                                    secondaryText: 'Latest updates on Azure products and features',
                                                    href: 'https://azure.microsoft.com/en-us/updates/',
                                                    target: '_blank'
                                                },
                                            }
                                        ]} />
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
        var dashboardMap: IRearrangeableGridDefinition | undefined = globalState.getDashboard("AzureHome");

        this.setState({
            map: dashboardMap ? dashboardMap : {
                azureActiveAlertsGraph: { cellHeight: 1, cellWidth: 1, row: 0, col: 0 },
                azureActiveAlertsList: { cellHeight: 1, cellWidth: 2, row: 0, col: 1 }
            }
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
        globalState.setDashboard("AzureHome", newMap);

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
                globalState.setDashboard("M365Home", newMap);

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
            globalState.setDashboard("AzureHome", this.state.map);
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