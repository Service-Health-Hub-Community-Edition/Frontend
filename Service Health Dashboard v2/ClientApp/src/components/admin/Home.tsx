import * as React from 'react';
import { Text } from '@fluentui/react';
import { AdminLayout } from "./AdminLayout";
import { Redirect } from 'react-router-dom';
import { Image } from '@fluentui/react';
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

interface IAdminHomeState {
    panelOpen: boolean;
    isLoading: boolean;
    useCompaction: boolean;
    map: IRearrangeableGridDefinition | undefined;
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
    large = 'Large',
}

type SizeMap = { [key in CardSizes]: { cellWidth: number; cellHeight: number } };

const CardSizeValues: SizeMap = {
    Small: { cellWidth: 1, cellHeight: 1 },
    'Medium (tall)': { cellWidth: 1, cellHeight: 2 },
    'Medium (wide)': { cellWidth: 2, cellHeight: 1 },
    Large: { cellWidth: 2, cellHeight: 2 },
};

const addCardPanelInfo: AddCardPanelInfoType = {
};

const removeText = 'Remove';
const resizeText = 'Resize';

export class AdminHome extends React.Component<{}, IAdminHomeState> {
    constructor(props: {}) {
        super(props);

        this.state = {
            panelOpen: false,
            isLoading: false,
            useCompaction: true,
            map: {
                notificationSettings: { cellHeight: 1, cellWidth: 2, row: 0, col: 0 },
                metadataMappingSettings: { cellHeight: 1, cellWidth: 1, row: 1, col: 0 },
                integrationSettings: { cellHeight: 1, cellWidth: 1, row: 1, col: 1 },
            },
            changeMessage: ""
        }

    }

    public render() {
        const { panelOpen, isLoading, useCompaction, map } = this.state;

        return (
            <AdminLayout>
                <div className="container" style={{ paddingTop: '15px', paddingBottom: '15px' }}>
                    <div className="row">
                        <div className="col">
                    <Dashboard
                        isAddCardPanelOpen={panelOpen}
                        map={map}
                        onChange={this._onChange}
                        onAddRemove={this._onAddRemove}
                        
                        ariaLabel="Microsoft Service Health Hub Admin Center"
                        ariaDescription="Use the arrow keys to navigate to a card. Press alt-shift-arrow to move a card. Press enter key to enter each item. Press escape to return to the grid when within an item."
                        rearrangeableGridProps={{

                            ...(useCompaction && { onCompaction: compactVerticalThenHorizontal }),
                        }}
                        panelProps={{
                            closeButtonAriaLabel: 'Close add card panel',
                            headerText: 'Drag cards to your home page',
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
                                src: '/images/addCardPanelVerticalBarEmptyState.png',
                                ['aria-hidden']: 'true',
                            },
                            illustrationSize: EmptyStateImageSize.Large,
                        }}
                    >

                        <DashboardCard
                            key="notificationSettings"
                            headerText={<><br />Configure notifications</>}
                            titleText="Notifications"
                            styles={{
                                secondaryBody: {
                                    // This effectively makes the right size a minimum width of 282,
                                    // before reflowing to a vertical orientation
                                    flex: '1 1 282px',
                                },
                            }}
                            body={
                                <DashboardCardBodyText
                                    subHeaderText={'Manage notification routing'}
                                    bodyText={
                                        'Define connections for alert distribution and define notification routing rules for Service Health Hub components.'
                                    }
                                />
                            }
                            secondaryBody={
                                <ThumbnailItemStack items={
                                    [
                                        {
                                            imageProps: { src: '/images/admin.svg', alt: 'Connectors' },
                                            compoundButtonProps: {
                                                primary: false,
                                                text: 'Connectors',
                                                secondaryText: 'Configure alert connectors',
                                                href: '/admin/notifications/connectors'
                                            }
                                        },
                                        {
                                            imageProps: { src: '/images/customize.svg', alt: 'Connectors' },
                                            compoundButtonProps: {
                                                primary: false,
                                                text: 'Routing',
                                                secondaryText: 'Configure alert routing',
                                                href: '/admin/notifications/entities'
                                            }
                                        }
                                    ]
                                } />
                            }

                            footer={
                                <ActionBar primaryButtonProps={{
                                    text: 'Get started',
                                    href: '/admin/notifications'
                                }} />
                            }
                        />

                        <DashboardCard
                            key="metadataMappingSettings"
                            headerText={<><br/>Metadata mapping</>}
                            titleText="Task metadata mapping"
                            styles={{
                                secondaryBody: {
                                    // This effectively makes the right size a minimum width of 282,
                                    // before reflowing to a vertical orientation
                                    flex: '1 1 282px',
                                },
                            }}
                            body={
                                <div>
                                    <DashboardCardBodyText
                                        subHeaderText={'Map component data'}
                                        bodyText={
                                            'Map Service Health Hub component metadata to task fields in Azure DevOps, ServiceNow or Jira.'
                                        }
                                    />
                                    <br/><br/>
                                    <div className="container" style={{ padding: '0px' }}>
                                        <div className="row">
                                            <div className="col ml-auto" style={{ marginLeft: '0px', marginRight: '0px' }} >
                                                <div className="float-right">
                                                    <Image src='/images/tasks.svg' style={{ opacity: '75%', height: '128px' }} />
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            }
                            footer={
                                <ActionBar primaryButtonProps={{
                                    text: 'Configure',
                                    href: '/admin/mapping'
                                }} />
                            }
                        />

                        <DashboardCard
                            key="integrationSettings"
                            headerText={<><br/>Integration</>}
                            titleText="Custom actions"
                            styles={{
                                secondaryBody: {
                                    // This effectively makes the right size a minimum width of 282,
                                    // before reflowing to a vertical orientation
                                    flex: '1 1 282px',
                                },
                            }}
                            body={
                                <div>
                                    <DashboardCardBodyText
                                        subHeaderText={'Integration settings'}
                                        bodyText={
                                            'Implement your business logic by connecting Microsoft Service Health Hub to your Power Automate Flows, Logic Apps, Azure Functions and other external systems.'
                                        }
                                    />
                                    <div className="container" style={{ padding: '0px' }}>
                                        <div className="row">
                                            <div className="col ml-auto" style={{ marginLeft: '0px', marginRight: '0px' }} >
                                                <div className="float-right">
                                                    <Image src='/images/logicApp.svg' style={{ opacity: '75%', height: '160px' }} />
                                                </div>
                                            </div>
                                        </div>
                                     </div>
                                </div>
                            }
                            footer={
                                <ActionBar primaryButtonProps={{
                                    text: 'Configure',
                                    href: '/admin/customActions/components'
                                }} />
                            }
                        />
                    </Dashboard>
                        </div>
                    </div>
                </div>

            </AdminLayout>
        );
    }

    private _openPanel = (): void => {

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
                // resizeCallback.current?.(dimension.cellWidth, dimension.cellHeight, key);
            }
        };
    };

    private _generateItemMenu = (key: string): IContextualMenuProps => {
        return {
            items: [
                /* {
                    iconProps: { iconName: 'Cancel' },
                    key: removeText,
                    text: removeText,
                    onClick: this._generateRemoveCallback(key),
                    ariaLabel: 'Remove card',
                }, */
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
                },
            ],
        };
    };

    private _onAddRemove = (
        additions: IAddCardPanelItemProps[],
        removals: IAddCardPanelItemProps[],
    ) => {
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