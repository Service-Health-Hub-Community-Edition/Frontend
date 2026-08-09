import * as React from 'react';
import {
    Label, Text, Spinner, SpinnerSize, mergeStyles, mergeStyleSets,
    DetailsList, DetailsListLayoutMode, IColumn, SelectionMode, FontIcon, TooltipHost,
    ITooltipHostStyles, Panel, PanelType, Link, ILinkStyles, ILinkStyleProps
} from '@fluentui/react';
import { acquireAccessToken } from "../../../auth/AccessTokenHelper";
import { AuthenticationResult } from '@azure/msal-browser';
import { JobDetailsList } from './JobDetailsList';
import { IAuditLogEntry } from '../../AuditLog';

export interface IJobHistory {
    correlationId: string;
    jobName: string;
    start?: Date;
    end?: Date;
    state: string;
    itemsCreated: number;
    itemsModified: number;
    itemsFailed: number;
    tasksCreated: number;
    tasksModified: number;
    tasksFailed: number;
    notificationsSent: number;
    notificationsFailed: number;
}

interface ISelectedLogEntry {
    correlationId: string;
    timestamp?: Date;
}

interface IJobHistoryState {
    componentId: string;
    communicationId?: string;
    maxItems?: number;
    selectedEntry?: ISelectedLogEntry;
    items: IJobHistory[];
    hideEmpty?: boolean;
    startDate?: Date;
    endDate?: Date;
    accessGranted: boolean;
    initialized: boolean;
    error?: string;
}

const iconClass = mergeStyles({
    margin: '0 0px',
});

const classNames = mergeStyleSets({
    red: [{ color: 'red' }, iconClass],
    darkgreen: [{ color: 'darkgreen' }, iconClass],
    green: [{ color: 'green' }, iconClass],
    blue: [{ color: '#007bff' }, iconClass],
    deepSkyBlue: [{ color: 'deepskyblue' }, iconClass],
    greenYellow: [{ color: 'greenyellow' }, iconClass],
    salmon: [{ color: 'salmon' }, iconClass],
});

export class JobHistory extends React.Component<{ componentId: string, maxItems?: number, communicationId?: string, hideEmpty?: boolean, startDate?: Date, endDate?: Date }, IJobHistoryState> {

    constructor(props: { componentId: string, maxItems?: number, communicationId?: string, hideEmpty?: boolean, startDate?: Date, endDate?: Date }) {
        super(props);

        this.state = {
            componentId: props.componentId,
            maxItems: props.maxItems,
            communicationId: props.communicationId,
            items: [],
            accessGranted: false,
            initialized: false
        }
    }

    render() {
        const { items, selectedEntry, accessGranted, initialized, error } = this.state;

        if (!initialized)
            return (<>
                <br /><br />
                <Spinner size={SpinnerSize.medium} />
            </>);

        if (error !== undefined)
            return (<>
                <br /><br />
                <Text variant='medium'><b>{error}</b></Text>
            </>);

        if (!accessGranted)
            return (<>
                <br /><br />
                <Text variant='medium'><b>You have no access to this ressource.</b></Text>
            </>);

        const calloutProps = { gapSpace: 0 };
        const hostStyles: Partial<ITooltipHostStyles> = { root: { display: 'inline-block' } };

        const pipeFabricStyles = (p: ILinkStyleProps): ILinkStyles => ({
            root: {
                textDecoration: 'none',
                color: p.theme.semanticColors.bodyText,
                fontWeight: '600',
                fontSize: p.theme.fonts.medium.fontSize,
            },
        });

        const columns: IColumn[] = [
            {
                key: '1',
                name: 'Time',
                fieldName: 'timestamp',
                minWidth: 120,
                isResizable: true,
                isRowHeader: true,
                onRender: (item: IJobHistory) => {
                        return <Link onClick={(event: any) => {
                            event.preventDefault();
                            this._onOpenDetailsPanel(item);
                        }} styles={pipeFabricStyles}>{item.start ? new Date(item.start).toLocaleString() : ""}</Link>
                }
            },
            {
                key: '2',
                name: 'Duration',
                fieldName: 'timestamp',
                minWidth: 160,
                maxWidth: 160,
                isResizable: false,
                onRender: (item: IJobHistory) => {
                    var duration: string = '---';
                    if (item.start) {
                        const durationms = (item.end ? new Date(item.end).valueOf() : new Date().valueOf()) - new Date(item.start).valueOf()
                        const durationDateTime = new Date(durationms);

                        duration = String(durationDateTime.getUTCHours()).padStart(2, "0") + ":" +
                            String(durationDateTime.getUTCMinutes()).padStart(2, "0") + ":" +
                            String(durationDateTime.getUTCSeconds()).padStart(2, "0")
                    }
                    return (<><Text>{duration}</Text></>)
                }
            },
            {
                key: '3',
                name: 'State',
                fieldName: 'state',
                isResizable: false,
                minWidth: 96,
                maxWidth: 96,
                onRender: (item: IJobHistory) => {
                    var iconName: string = 'CompletedSolid';
                    var iconColor: any = classNames.salmon;

                    switch (item.state) {
                        case 'Running':
                            iconName = 'SyncStatusSolid';
                            iconColor = classNames.blue;
                            break;
                        case 'Completed':
                            iconName = 'CompletedSolid';
                            iconColor = classNames.green;
                            break;
                        case 'Failed':
                            iconName = 'StatusErrorFull';
                            iconColor = classNames.red;
                            break;
                        case 'Timed out':
                            iconName = 'StatusErrorFull';
                            iconColor = classNames.red;
                            break;
                        default:
                            iconName = 'InfoSolid';
                            iconColor = classNames.greenYellow;
                    }
                    return (<div style={{ marginRight: '8px' }}>
                        <FontIcon iconName={iconName} className={iconColor} />&nbsp;&nbsp;{item.state}
                    </div>)
                }
            },
            {
                key: '4',
                name: 'Items',
                minWidth: 192,
                maxWidth: 192,
                isResizable: false,
                onRender: (item: IJobHistory) => {
                    return (<>
                        {item.itemsCreated === 0 && item.itemsModified === 0 && item.itemsFailed === 0 ? '---' : ''}

                        {item.itemsCreated > 0 ? (<>
                            <TooltipHost
                                content="Items created"
                                id={'item-created-' + item.correlationId}
                                calloutProps={calloutProps}
                                styles={hostStyles}
                            >
                                <FontIcon iconName='TextDocument' className={classNames.blue} />&nbsp;&nbsp;{item.itemsCreated}&nbsp;&nbsp;&nbsp;&nbsp;
                            </TooltipHost>
                        </>) : ''}

                        {item.itemsModified > 0 ? (<>
                            <TooltipHost
                                content="Items updated"
                                id={'item-updated-' + item.correlationId}
                                calloutProps={calloutProps}
                                styles={hostStyles}
                            >
                                <FontIcon iconName='Edit' className={classNames.blue} />&nbsp;&nbsp;{item.itemsModified}&nbsp;&nbsp;&nbsp;&nbsp;
                            </TooltipHost>
                        </>) : ''}

                        {item.itemsFailed > 0 ? (
                            <>
                                <TooltipHost
                                    content="Items failed"
                                    id={'item-failed-' + item.correlationId}
                                    calloutProps={calloutProps}
                                    styles={hostStyles}
                                >
                                    <FontIcon iconName='StatusErrorFull' className={classNames.red} />&nbsp;&nbsp;{item.itemsFailed}
                                </TooltipHost>
                            </>) : ''}
                    </>)
                }
            },
            {
                key: '5',
                name: 'Tasks',
                minWidth: 192,
                maxWidth: 192,
                isResizable: false,
                onRender: (item: IJobHistory) => {
                    return (<>
                        {item.tasksCreated === 0 && item.tasksModified === 0 && item.tasksFailed === 0 ? '---' : ''}

                        {item.tasksCreated > 0 ? (<>
                            <TooltipHost
                                content="Tasks created"
                                id={'task-created-'+item.correlationId}
                                calloutProps={calloutProps}
                                styles={hostStyles}
                            >
                                <FontIcon iconName='WorkItem' className={classNames.blue} />&nbsp;&nbsp;{item.tasksCreated}&nbsp;&nbsp;&nbsp;&nbsp;
                            </TooltipHost>
                        </>) : ''}

                        {item.tasksModified > 0 ? (<>
                            <TooltipHost
                                content="Tasks updated"
                                id={'task-updated-' + item.correlationId}
                                calloutProps={calloutProps}
                                styles={hostStyles}
                            >
                                <FontIcon iconName='Edit' className={classNames.blue} />&nbsp;&nbsp;{item.tasksModified}&nbsp;&nbsp;&nbsp;&nbsp;
                            </TooltipHost>
                        </>) : ''}

                        {item.tasksFailed > 0 ? (
                            <>
                                <TooltipHost
                                    content="Tasks failed"
                                    id={'task-failed-' + item.correlationId}
                                    calloutProps={calloutProps}
                                    styles={hostStyles}
                                >
                                    <FontIcon iconName='StatusErrorFull' className={classNames.red} />&nbsp;&nbsp;{item.tasksFailed}
                                </TooltipHost>
                            </>) : ''}
                    </>)
                }
            },
            {
                key: '6',
                name: 'Notifications',
                minWidth: 128,
                maxWidth: 128,
                isResizable: false,
                onRender: (item: IJobHistory) => {
                    return (<div>
                        {item.notificationsSent === 0 && item.notificationsFailed === 0 ? '---' : ''}
                        {item.notificationsSent > 0 ? (<>
                            <TooltipHost
                                content="Notifications sent"
                                id={'notification-sent-' + item.correlationId}
                                calloutProps={calloutProps}
                                styles={hostStyles}
                            >
                                <FontIcon iconName='CompletedSolid' className={classNames.green} />&nbsp;&nbsp;{item.notificationsSent}&nbsp;&nbsp;&nbsp;&nbsp;
                            </TooltipHost>
                        </>) : ''}

                        {item.notificationsFailed > 0 ? (
                            <>
                                <TooltipHost
                                    content="Notifications failed"
                                    id={'notification-failed-' + item.correlationId}
                                    calloutProps={calloutProps}
                                    styles={hostStyles}
                                >
                                    <FontIcon iconName='StatusErrorFull' className={classNames.red} />&nbsp;&nbsp;{item.notificationsFailed}
                                </TooltipHost>
                            </>) : ''}
                    </div>)
                }
            }
        ];

        return (<>
            <DetailsList
                selectionMode={SelectionMode.none}
                layoutMode={DetailsListLayoutMode.justified}
                compact={true}
                items={items}
                columns={columns}
                onShouldVirtualize={() => false}
            />
            <Panel
                headerText={'Job details' + (selectedEntry?.timestamp ? ': ' + new Date(selectedEntry.timestamp).toLocaleString() : '')}
                isOpen={selectedEntry !== undefined}
                onDismiss={() => this._onDismissViewDetailsPanel()}
                type={PanelType.large}
                isBlocking={true}
                allowTouchBodyScroll={true}
                closeButtonAriaLabel="Close"
            >
                {selectedEntry ? (<JobDetailsList id={selectedEntry.correlationId} />) : ""}
            </Panel>
        </>);
    }

    componentDidMount() {
        this.setState({
            componentId: this.props.componentId,
            maxItems: this.props.maxItems,
            communicationId: this.props.communicationId,
            hideEmpty: this.props.hideEmpty,
            startDate: this.props.startDate,
            endDate: this.props.endDate
        });

        this._loadData();
    }

    componentDidUpdate(prevProps: any) {
        
        if (prevProps.componentId !== this.props.componentId ||
            prevProps.maxItems !== this.props.maxItems ||
            prevProps.communicationId !== this.props.communicationId ||
            prevProps.hideEmpty !== this.props.hideEmpty ||
            prevProps.startDate !== this.props.startDate ||
            prevProps.endDate !== this.props.endDate) {

            this.setState({
                componentId: this.props.componentId,
                maxItems: this.props.maxItems,
                communicationId: this.props.communicationId,
                hideEmpty: this.props.hideEmpty,
                startDate: this.props.startDate,
                endDate: this.props.endDate
            });

            this._loadData();
        }
    }

    _loadData(): void {
        const requiredRoles: string[] = ['Admin'];
        var authResponse: AuthenticationResult;
        var userHasRequiredRole: boolean = false;

        var queryParams: string[] = [];
        if (this.props.maxItems)
            queryParams.push('top=' + this.props.maxItems);

        if (this.props.communicationId)
            queryParams.push('communicationId=' + this.props.communicationId);

        if (this.props.hideEmpty)
            queryParams.push('hideEmptyJobs=' + this.props.hideEmpty);

        if (this.props.startDate)
            queryParams.push('startDate=' + this.props.startDate);

        if (this.props.endDate)
            queryParams.push('endDate=' + this.props.endDate);

        var url: string = '/api/jobs/' + this.props.componentId + '/history';

        if (queryParams.length > 0)
            url += '/?' + queryParams.join('&');

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
            })
            .then(() => {
                if (userHasRequiredRole)
                    fetch(url, { headers: authResponse.idToken === "" ? {} : { 'Authorization': `Bearer ${authResponse.idToken}` } })
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
                    .then(response => {
                        this.setState({
                            items: response,
                            initialized: true
                        });
                    })
            }).catch((err) => {
                this.setState({
                    error: err.message
                });
            });
    }

    _onOpenDetailsPanel(item: IJobHistory): void {
        this.setState({
            selectedEntry: {
                correlationId: item.correlationId,
                timestamp: item.start
            }
        });
    }

    _onDismissViewDetailsPanel(): void {
        this.setState({
            selectedEntry: undefined
        });
    }
}