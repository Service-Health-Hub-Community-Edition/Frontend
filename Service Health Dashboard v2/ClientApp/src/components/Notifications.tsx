import * as React from 'react';
import { GlobalState } from './GlobalState';
import { ActionCard, ActionCardMode } from '@m365-admin/card';
import {
    Fabric, Text, FontIcon, mergeStyles, mergeStyleSets,
    IColumn, DetailsList, SelectionMode, DetailsListLayoutMode,
    Panel, PanelType, CommandBar, Image, Overlay, Link
} from '@fluentui/react';

export enum NotificationState {
    Information,
    Warning,
    Error,
    Running,
    Success
}

export interface INotification {
    id: string,
    state: NotificationState,
    title: string,
    message: string,
    time: Date,
    active: boolean,
    onClick?(): void
}

export interface INotificationState {
    notifications: INotification[];
    columns: IColumn[];
}

interface IToastNotificationState {
    width: number
}

export class ToastNotification extends React.Component<{}, IToastNotificationState> {
    static contextType = GlobalState;

    constructor(props: {}) {
        super(props);

        this.state = {
            width: window.innerWidth
        }

        window.addEventListener('resize', () => this.handleResize());
    }

    componentDidMount() {
     
    }

    public render() {
        const { width } = this.state;
        let globalState = this.context;

        return globalState.toastNotification === undefined ? (<div />) : (
            <div style={{
                width: '300px',
                position: 'fixed',
                top: '56px',
                left: width - 310,
                zIndex: 10000000
            }}>
                <Overlay styles={{ root: { width: '300px' } }}>
                    <ActionCard
                        primaryText={globalState.toastNotification.title}
                        secondaryText={globalState.toastNotification.message}
                        color="dodgerblue"
                        mode={ActionCardMode.SecondaryPage}
                        styles={{ root: { width: '300px', boxShadow: '5px 5px 10px rgba(0, 0, 0, 0.2)' } }}
                        iconButtonProps={{
                            iconProps: {
                                iconName: 'ChromeClose',
                                styles: {
                                    root: {
                                        fontSize: 8,
                                        height: 8,
                                        width: 8,                                    }
                                }
                            },
                            title: 'Close',
                            height: 8,
                            width: 8, 
                            onClick: () => this.hideNotification()
                        }}
                    />
                </Overlay>
            </div>
        );
    }

    private hideNotification(): void {
        let globalState = this.context;
        globalState.clearToastNotification();
    }

    private handleResize(): void {
        this.setState({
            width: window.innerWidth
        });
    }
}


export class Notifications extends React.Component<{}, INotificationState> {
    static contextType = GlobalState;

    constructor(props: {}) {
        super(props);

        const iconClass = mergeStyles({
            fontSize: 16,
            height: 16,
            width: 16,
            margin: "0 16px"
        });

        const classNames = mergeStyleSets({
            error: [{ color: '#dc3545' }, iconClass],
            ok: [{ color: '#107c10' }, iconClass],
            information: [{ color: 'dodgerblue' }, iconClass],
            warning: [{ color: '#ffb900' }, iconClass],
            sync: [{ color: 'dodgerblue' }, iconClass]
        });

        const iconList: any = {
            Failed: 'StatusErrorFull',
            Running: 'SyncStatusSolid',
            Success: 'CompletedSolid',
            Warning: 'WarningSolid',
            Information: 'InfoSolid'
        }

        const columns: IColumn[] = [
            {
                key: 'clEventType',
                name: 'Event type',
                fieldName: 'eventType',
                minWidth: 24,
                maxWidth: 24,
                isRowHeader: false,
                isResizable: false,
                onRender: (item: INotification) => {
                    var iconName: string = '';
                    var iconClass: any = classNames.information;
                    var ariaLabel: string = item.state ? item.state.toString() : 'Information';

                    switch (item.state) {
                        case NotificationState.Error:
                            iconName = iconList.Failed;
                            iconClass = classNames.error;
                            break;

                        case NotificationState.Running:
                            iconName = iconList.Running;
                            iconClass = classNames.sync;
                            break;

                        case NotificationState.Success:
                            iconName = iconList.Success;
                            iconClass = classNames.ok;
                            break;

                        case NotificationState.Warning:
                            iconName = iconList.Warning;
                            iconClass = classNames.warning;
                            break;

                        default:
                            iconName = iconList.Information;
                            iconClass = classNames.information;
                    }

                    return (<FontIcon aria-label={ariaLabel} iconName={iconName} className={iconClass} />)
                },
                data: 'string',
                isPadded: false,
            },
            {
                key: 'clTitle',
                name: 'Title',
                fieldName: 'title',
                minWidth: 200,
                isRowHeader: false,
                isResizable: false,
                isMultiline: true,
                onRender: (item: INotification) => {
                    return (<div>
                        <Text variant='medium'>
                            <b>{item.title}</b>
                        </Text><br/>
                        <Text variant='medium'>
                            <div dangerouslySetInnerHTML={{ __html: item.message }} />
                        </Text><br/>
                        <div className='content'>
                            <div className='row'>
                                <div className='col'>
                                    <Text variant='small'>{item.time.toLocaleDateString([], { year: 'numeric', month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' })}</Text>
                                </div>
                                <div className='col-auto' style={{ textAlign: 'right' }} >
                                    <Link onClick={() => this._RemoveNotification(item.id)} >Dismiss</Link>
                                </div>
                            </div>
                        </div>
                    </div>);
                },
                data: 'string',
                isPadded: false,
            }
        ];

        this.state = {
            notifications: [],
            columns: columns
        };
    }

    componentDidMount() {
        let globalState: any = this.context;
        var notifications: INotification[] = globalState.notifications;

        this.setState({
            notifications: notifications
        });
    }

    public render() {
        const { columns } = this.state;
        let globalState: any = this.context;

        return (
            <div>
            <Panel
                headerText='Notifications'
                isOpen={globalState.isNotificationsPanelOpen}
                onDismiss={() => this._onDismissPanel()}
                type={PanelType.smallFixedFar}
                isLightDismiss
                closeButtonAriaLabel="Close">

                {globalState.notifications && globalState.notifications.length > 0 ? (
                    <div>
                        <CommandBar
                            items={[]}
                            farItems={[
                                {
                                    key: 'dismissAll',
                                    text: 'Dismiss all',
                                    onClick: () => { this._clearNotifications() }
                                }
                            ]}
                            />

                        <DetailsList
                            items={globalState.notifications}
                            compact={false}
                            columns={columns}
                            selectionMode={SelectionMode.none}
                            setKey="id"
                            layoutMode={DetailsListLayoutMode.justified}
                            isHeaderVisible={false}
                            isSelectedOnFocus={true}
                            />
                        </div>
                ) : (
                        <div style={{ textAlign: 'center' }} >
                            &nbsp;<br /><br/><br/>
                            <Image
                                src='/images/emptyNotificationPanel.svg'
                                style={{
                                    width: '100%',
                                    height: '100px',
                                    filter: 'grayscale(100%)',
                                    alignContent: 'center', textAlign: 'center'
                                }} />
                            <br /><br/>
                            <Text variant='medium'>
                                <b>No new notifications from this session</b>
                            </Text>
                    </div>)}
                </Panel>
                <ToastNotification/>
                </div>
        );
    }

    private _onDismissPanel(): void {
        let globalState: any = this.context;
        globalState.openNotificationsPanel(false);
    }

    private _clearNotifications(): void {
        let globalState: any = this.context;
        globalState.clearNotifications();
        this._onDismissPanel();
    }

    private _RemoveNotification(id: string): void {
        let globalState: any = this.context;
        globalState.removeNotification(id);
    }
}