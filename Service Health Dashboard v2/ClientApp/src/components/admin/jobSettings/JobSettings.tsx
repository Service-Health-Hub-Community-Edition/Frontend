import * as React from 'react';
import { Text } from '@fluentui/react';
import { AdminLayout } from "../AdminLayout";
import { Redirect } from 'react-router-dom';
import { Image, Label, Toggle } from '@fluentui/react';
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
import { EmptyStateImageSize } from '@m365-admin/empty-state';
import { ActionBar } from '@m365-admin/action-bar';
import { DetailPageHeader } from '@m365-admin/detail-page';
import { M365Breadcrumb } from '@m365-admin/m365-breadcrumb';
import { acquireAccessToken } from "../../../auth/AccessTokenHelper";
import { AuthenticationResult } from '@azure/msal-browser';
import { IComponent } from '../notifications/Routing';
import { JobSettingsEditor } from './JobSettingsEditor';
import { JobHistory } from './JobHistory';
import { JobStatistics } from './JobStatistics';

interface IJobSettingsState {
    componentId: string;
    componentName: string;
    panelOpen: boolean;
    isLoading: boolean;
    useCompaction: boolean;
    map: IRearrangeableGridDefinition | undefined;
    changeMessage: string;
    accessGranted: boolean;
    initialized: boolean;
    error?: string
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

export class JobSettings extends React.Component<{}, IJobSettingsState> {
    constructor(props: {}) {
        super(props);

        const queryParams = new URLSearchParams(window.location.search);
        const componentId = queryParams.get('id');

        this.state = {
            componentId: componentId && componentId !== null ? componentId : "",
            componentName: "",
            panelOpen: false,
            isLoading: false,
            useCompaction: true,
            map: {
                jobConfiguration: { cellHeight: 1, cellWidth: 1, row: 0, col: 0 },
                jobStatistics: { cellHeight: 1, cellWidth: 1, row: 0, col: 1 },
                jobHistory: { cellHeight: 1, cellWidth: 2, row: 1, col: 0 },
            },
            changeMessage: "",
            accessGranted: false,
            initialized: false
        }

    }

    public render() {
        const { panelOpen, isLoading, useCompaction, map, componentId, componentName } = this.state;

        return (
            <AdminLayout>
                <div className="container" style={{ maxWidth: '97%' }}>
                    <div className="row">
                        <div className="col">
                            <M365Breadcrumb
                                items={[
                                    { text: 'Home', key: 'home', href: '/admin' },
                                    { text: 'Jobs', key: 'jobs', href: '/admin/jobs' },
                                    { text: componentName, key: componentName, isCurrentItem: true }
                                ]}
                                style={{ marginBottom: '16px' }}
                            />
                            <DetailPageHeader
                                title={ componentName }
                                description="Edit configuration and view job statistics."
                            />
                        </div>
                    </div>
                    <div className="row">
                        <div className="col">
                    <Dashboard
                        isAddCardPanelOpen={panelOpen}
                        map={map}
                        onChange={this._onChange}
                        onAddRemove={this._onAddRemove}
                        
                        ariaLabel="Job settings"
                        ariaDescription="Use the arrow keys to navigate to a card. Press alt-shift-arrow to move a card. Press enter key to enter each item. Press escape to return to the grid when within an item."
                        rearrangeableGridProps={{

                            ...(useCompaction && { onCompaction: compactVerticalThenHorizontal })
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
                                src: '/images/addCardPanelVerticalBarEmptyState.png',
                                ['aria-hidden']: 'true',
                            },
                            illustrationSize: EmptyStateImageSize.Large,
                        }}
                    >

                        <DashboardCard
                            key="jobConfiguration"
                            headerText={<><br/>Job configuration</>}
                            titleText="Configuration"
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
                                        bodyText={
                                            'Configure ' + componentName + ' job. New settings will apply to the new job instances and won\'t affect currently running jobs.'
                                        }
                                    />
                                    <br/><br/>
                                    <div className="container" style={{ padding: '0px' }}>
                                        <div className="row">
                                            <div className="col" >
                                                <JobSettingsEditor componentId={componentId} />
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            }
                        />

                        <DashboardCard
                            key="jobStatistics"
                            headerText={<><br/>Job statistics</>}
                            titleText="Statistics"
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
                                        bodyText={
                                            'Job statistics for the past 24 hours.'
                                        }
                                    />
                                    <div className="container" style={{ padding: '0px' }}>
                                        <div className="row">
                                            <div className="col">
                                                <br/>
                                                <JobStatistics componentId={componentId} />
                                            </div>
                                        </div>
                                     </div>
                                </div>
                            }
                            /*footer={
                                <ActionBar primaryButtonProps={{
                                    text: 'More',
                                    href: '/admin/customActions/components'
                                }} />
                            }*/
                                />

                                <DashboardCard
                                    key="jobHistory"
                                    titleText="History"
                                    styles={{
                                        secondaryBody: {
                                            // This effectively makes the right size a minimum width of 282,
                                            // before reflowing to a vertical orientation
                                            flex: '1 1 282px',
                                        },
                                    }}
                                    body={
                                        <JobHistory componentId={componentId} maxItems={5} />
                                    }

                                    footer={
                                        <ActionBar primaryButtonProps={{
                                            text: 'See all',
                                            href: '/admin/jobs/history?id='+componentId
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

    componentDidMount() {
        this._loadData();
    }

    _loadData() {
        const requiredRoles: string[] = ['Admin'];
        var authResponse: AuthenticationResult;
        var userHasRequiredRole: boolean = false;
        var component: IComponent | undefined = undefined;

        this.setState({
            initialized: false
        });

        acquireAccessToken()
            .then((response) => {
                var tokenClaims: any = response.account?.idTokenClaims;
                const userRoles: any = tokenClaims?.roles;
                userHasRequiredRole = userRoles.some((r: string) => requiredRoles.includes(r));

                this.setState({
                    accessGranted: userHasRequiredRole
                });

                authResponse = response;
            }).then(() => {
                if (userHasRequiredRole)
                    fetch('/api/Components?id=' + this.state.componentId, { headers: authResponse.idToken === "" ? {} : { 'Authorization': `Bearer ${authResponse.idToken}` } })
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
                            if (result !== undefined && result.length > 0)
                                component = result[0] as IComponent;

                            this.setState({
                                componentName: component ? component?.name : ""
                            });
                        }).catch((err) => {
                            this.setState({
                                error: err.message
                            });
                        });
            });
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