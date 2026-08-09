import * as React from 'react';
import { ILinkStyleProps, ILinkStyles, Link, Text, FontIcon } from '@fluentui/react';
import { Fabric, mergeStyles, mergeStyleSets } from '@fluentui/react';
import { DetailsList, DetailsListLayoutMode, Selection, SelectionMode, IColumn, IDetailsListStyles } from '@fluentui/react';
import { Dialog, DialogType, DialogFooter, PrimaryButton } from '@fluentui/react';
import { DiffViewer } from './DiffViewer';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { MessageBar, MessageBarType } from '@fluentui/react';
import { acquireAccessToken } from "../auth/AccessTokenHelper";
import { checkInTeams } from './auth/detectTeams';
import { AccessDenied } from "./AccessDenied";

export interface IAuditLogState {
    itemId: string | undefined;
    itemType: string | undefined;
    scheme: string;
    columns: IColumn[];
    items: IAuditLogEntry[];
    selectedAuditLogEntry: IAuditLogEntry | undefined;
    inTeams: boolean;
    isLoaded: boolean;
    error?: string;
    accessGranted?: boolean;
}

export interface IAuditLogEntry {
    key: string;
    id: string;
    timestamp: Date;
    user: string | undefined;
    activity: string;
    item: string;
    itemType: string;
    itemUniqueId: string;
    eventSource: string;
    extendedProperties: any;
    modifiedProperties: any;
}

export class AuditLog extends React.Component<{ itemId: string, itemType: string, scheme: string | undefined }, IAuditLogState> {

    constructor(props: { itemId: string, itemType: string, scheme: string | undefined }) {
        super(props);

        const pipeFabricStyles = (p: ILinkStyleProps): ILinkStyles => ({
            root: {
                textDecoration: 'none',
                color: p.theme.semanticColors.bodyText,
                fontWeight: '600',
                fontSize: p.theme.fonts.medium.fontSize,
            },
        });

        const iconClass = mergeStyles({
            fontSize: 16,
            height: 16,
            width: 16,
            margin: "0 16px"
        });

        const classNames = mergeStyleSets({
            error: [{ color: '#dc3545' }, iconClass],
            deepSkyBlue: [{ color: 'deepskyblue' }, iconClass],
            greenYellow: [{ color: 'greenyellow' }, iconClass],
            salmon: [{ color: 'salmon' }, iconClass],
            black: [{ color: 'black' }, iconClass],
        });

        const iconList: any = {
            ActionSuccess: 'TriggerUser',
            Task: 'TaskSolid',
            TaskCreated: 'TaskSolid',
            TaskModified: 'Edit',
            Item: 'Breadcrumb',
            Message: 'Breadcrumb',
            Notification: 'Ringer',
            Created: 'TextDocument',
            Modified: 'Edit',
            Published: 'PublishContent',
            Unpublished: 'UnpublishContent',
            Archived: 'Archive',
            Restored: 'Undo',
            Failed: 'StatusErrorFull',
            Search: 'Search',
            default: 'Breadcrumb'
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
                onRender: (item: IAuditLogEntry) => {
                    var iconName: string = '';
                    var iconClass: any = classNames.black;

                    switch (item.itemType.toLowerCase()) {                    
                        case 'message':
                        case 'item':
                            switch (item.activity.toLowerCase()) {
                                case 'itemadded':
                                case 'created':
                                    iconName = iconList.Created;
                                    break;

                                case 'itemmodified':
                                case 'modified':
                                    iconName = iconList.Modified;
                                    break;

                                case 'published':
                                    iconName = iconList.Published;
                                    break;

                                case 'unpublished':
                                    iconName = iconList.Unpublished;
                                    break;

                                case 'archived':
                                    iconName = iconList.Archived;
                                    break;

                                case 'restored':
                                    iconName = iconList.Restored;
                                    break;

                                case 'indexsuccess':
                                    iconName = iconList.Search;
                                    break;

                                case 'indexfailed':
                                    iconName = iconList.Failed;
                                    iconClass = classNames.error;
                                    break;

                                default:
                                    if (item.activity.toLowerCase().indexOf('failed') > -1) {
                                        iconName = iconList.Failed;
                                        iconClass = classNames.error;
                                    } else
                                        iconName = iconList.default;
                            }
                            break;

                        case 'action':
                            switch (item.activity.toLowerCase()) {
                                case 'success':
                                    iconName = iconList.ActionSuccess;
                                    break;

                                case 'failed':
                                    iconName = iconList.Failed;
                                    iconClass = classNames.error;
                                    break;

                                default:
                                    if (item.activity.toLowerCase().indexOf('failed') > -1) {
                                        iconName = iconList.Failed;
                                        iconClass = classNames.error;
                                    } else
                                        iconName = iconList.default;
                            }
                            break;

                        case 'notification':
                            iconName = iconList.Notification;
                            break;

                        case 'task':
                            switch (item.activity.toLowerCase()) {
                                case 'created':
                                    iconName = iconList.TaskCreated;
                                    break;

                                case 'modified':
                                    iconName = iconList.TaskModified;
                                    break;

                                case 'createfailed':
                                    iconName = iconList.Failed;
                                    iconClass = classNames.error;
                                    break;

                                case 'updatefailed':
                                    iconName = iconList.Failed;
                                    iconClass = classNames.error;
                                    break;

                                default:
                                    iconName = iconList.default;
                                    break;
                            }

                            break;

                        default:
                            iconName = iconList.default;
                    }

                    return (<FontIcon aria-label={item.itemType} iconName={iconName} className={iconClass} />)
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
                onRender: (item: IAuditLogEntry) => {
                    var message: string = '';
                    var addViewChanges: boolean = false;

                    switch (item.itemType.toLowerCase()) {
                        case 'message':
                        case 'item':
                            switch (item.activity.toLowerCase()) {
                                case 'itemadded':
                                case 'created':
                                    message = 'Item created';
                                    if (item.eventSource.toLowerCase() === 'databaseupgrade')
                                        message += ' through a version upgrade';
                                    break;

                                case 'itemmodified':
                                case 'modified':
                                    message = 'Item modified';
                                    if (item.modifiedProperties) {
                                        addViewChanges = true;
                                        message += '. ';
                                    }
                                    break;

                                case 'published':
                                    message = 'Item published';
                                    break;

                                case 'unpublished':
                                    message = 'Item unpublished';
                                    break;

                                case 'archived':
                                    message = 'Item archived';
                                    break;

                                case 'restored':
                                    message = 'Item restored';
                                    break;

                                case 'indexsuccess':
                                    message = 'Item indexed';
                                    break;

                                case 'indexfailed':
                                    message = 'Indexing failed';
                                    break;

                                default:
                                    message = 'Operation performed on a communication item. Activity code: ' + item.activity.toLowerCase();
                                    if (item.eventSource.toLowerCase() === 'databaseupgrade')
                                        message += '. Source: Database upgrade';
                            }

                            break;
                        case 'action':
                            switch (item.activity.toLowerCase()) {
                                case 'success':
                                    message = 'Action <b>' + item.extendedProperties.actionName + '</b> triggered. Result: ' + item.extendedProperties.actionResponse.ResponseBody;
                                    break;

                                case 'failed':
                                    message = 'Action <b>' + item.extendedProperties.actionName + '</b> failed. Result: ' + item.extendedProperties.actionResponse.ResponseBody;
                                    break;

                                default:
                                    message = 'Operation performed on a communication item. Activity code: ' + item.activity.toLowerCase();
                            }

                            break;
                        case 'notification':
                            message = item.extendedProperties.Route ?
                                (Array.isArray(item.extendedProperties.Route) && (item.extendedProperties.Route.length > 1) ? 
                                    'Notifications sent to <ul>' + item.extendedProperties.Route.map((route: any) => '<li style="font-style: italic; font-size: 12px">' + route + '</li>').join(' ') + '</ul>' : 
                                    'Notification sent to <i>' + item.extendedProperties.Route + '</i>') : 'Notification sent';

                            break;
                        case 'task':
                            switch (item.activity.toLowerCase()) {
                                case 'created':
                                    message =
                                        item.extendedProperties.TaskUrl ?
                                            'Task <a href="' + item.extendedProperties.TaskUrl + '" target="_blank">' + item.extendedProperties.TaskId + '</a> created in ' + item.extendedProperties.TaskManager :
                                            'Task created' + item.extendedProperties.TaskManager ? ' in ' + item.extendedProperties.TaskManager : '';
                                    break;

                                case 'modified':
                                    message =
                                        item.extendedProperties.TaskUrl ?
                                            'Task <a href="' + item.extendedProperties.TaskUrl + '" target="_blank">' + item.extendedProperties.TaskId + '</a> modified in ' + item.extendedProperties.TaskManager :
                                            'Task created' + item.extendedProperties.TaskManager ? ' in ' + item.extendedProperties.TaskManager : '';
                                    break;

                                case 'createfailed':
                                    message =
                                        item.extendedProperties.TaskManager ?
                                            'Failed to create task in ' + item.extendedProperties.TaskManager :
                                            'Failed to create task';
                                    break;

                                case 'updatefailed':
                                    message =
                                        item.extendedProperties.TaskManager ?
                                            'Failed to modify task in ' + item.extendedProperties.TaskManager :
                                            'Failed to modify task';
                                    break;

                                default:
                                    message =
                                        item.extendedProperties.TaskUrl ?
                                            'Operation performed on task <a href="' + item.extendedProperties.TaskUrl + '" target="_blank">' + item.extendedProperties.TaskId + '</a> in ' + item.extendedProperties.TaskManager + '. Activity code: ' + item.activity :
                                            'Operation performed on task' + (item.extendedProperties.TaskManager ? ' in ' + item.extendedProperties.TaskManager : '') + '. Activity code: ' + item.activity;
                                    break;
                            }

                            break;

                        default:
                    }
                    return (<div>
                        <Text variant='medium'>
                            <div dangerouslySetInnerHTML={{ __html: message }} />
                        </Text>
                        <div className='content'>
                            <div className='row'>
                                <div className='col'>
                                    <Text variant='small'>{item.timestamp.toLocaleDateString([], { year: 'numeric', month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' })}</Text>
                                </div>
                                {addViewChanges ? (
                                    <div className='col' style={{ textAlign: 'right' }} >
                                        <Link onClick={() => this._openViewChangesDialog(item)} >View changes</Link>
                                    </div>) : (<div />)}
                            </div>
                        </div>
                    </div>);
                },
                data: 'string',
                isPadded: false,
            }
        ];

        this.state = {
            itemId: this.props.itemId,
            itemType: this.props.itemType,
            scheme: this.props.scheme ? this.props.scheme : 'item',
            items: [],
            selectedAuditLogEntry: undefined,
            columns: columns,
            inTeams: checkInTeams(),
            isLoaded: false,
            error: undefined,
            accessGranted: undefined
        };
    }

    public render() {
        const {
            columns, items, selectedAuditLogEntry, isLoaded, error, accessGranted
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

        if (!isLoaded)
            return (
                <div className="loadingProgress">
                    <br /><br />
                    <Spinner size={SpinnerSize.large} />
                </div>
            );

        if (error)
            return (
                <div>
                    <MessageBar
                        messageBarType={MessageBarType.error}
                        isMultiline={false}
                    >
                        Couldn't retrieve data. Error: {error}
                    </MessageBar>
                    <br />
                </div>
                );

        return (
            <Fabric>
                <div className="container">
                    <div className="row">
                        <div className="col">
                            {items && items.length > 0 ? (
                                <DetailsList
                                    items={items}
                                    compact={false}
                                    columns={columns}
                                    selectionMode={SelectionMode.none}
                                    getKey={this._getKey}
                                    setKey="id"
                                    layoutMode={DetailsListLayoutMode.justified}
                                    isHeaderVisible={false}
                                    onItemInvoked={this._onItemInvoked}
                                    isSelectedOnFocus={true}
                                />
                            ) : (<div style={{ textAlign: 'center'}} >&nbsp;<br/><Text variant='medium'>No activities available</Text></div>)}                          
                        </div>
                    </div>
                </div>

                {selectedAuditLogEntry !== undefined ? (
                    <Dialog
                        hidden={false}
                        onDismiss={() => this._closeViewChangesDialog(this)}
                        dialogContentProps={{
                            type: DialogType.normal,
                            title: selectedAuditLogEntry.itemUniqueId + ': modified properties',
                        }}
                        modalProps={{
                            isBlocking: true
                        }}
                        styles={{
                            main: {
                                selectors: {
                                    ['@media (min-width: 480px)']: {
                                        width: 640,
                                        minWidth: 300,
                                        maxWidth: '1000px'
                                    }
                                }
                            }
                        }}
                    >
                        <DiffViewer diffObject={selectedAuditLogEntry.modifiedProperties} />

                        <DialogFooter>
                            <PrimaryButton onClick={() => this._closeViewChangesDialog(this)} text="Close" />
                        </DialogFooter>
                    </Dialog>
                ) : (<div />)}
            </Fabric>
        );
    }

    public componentDidUpdate(previousProps: any, previousState: IAuditLogState) {
        if (previousProps.itemId !== this.props.itemId || previousProps.itemType !== this.props.itemType)
            this._onLoadData();
    }

    componentDidMount() {
        this._onLoadData();
    }

    _onLoadData() {

        this.setState({
            isLoaded: false
        });

        var items: IAuditLogEntry[] = [];
        var processedIds: string[] = [];
        const requiredRoles: string[] = ['ServiceHealthReader', 'Communication.Write.All', 'Admin'];

        acquireAccessToken()
            .then((response) => {
                var tokenClaims: any = response.account?.idTokenClaims;
                const userRoles: any = tokenClaims?.roles;
                var userHasRequiredRole: boolean = userRoles.some((r: string) => requiredRoles.includes(r));

                this.setState({
                    accessGranted: userHasRequiredRole
                });

                if (userHasRequiredRole)
                    fetch('/api/AuditLog?scheme=' + this.state.scheme + '&id=' + this.state.itemId + '&type=' + this.state.itemType, { headers: response.idToken === "" ? {} : { 'Authorization': `Bearer ${response.idToken}` } })
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
                        for (const logEntry of result) {
                            if (processedIds.find(i => i === logEntry.id) === undefined) {
                                if (logEntry.itemType.toLowerCase() === 'notification'
                                    && !(logEntry.correlationId === undefined || logEntry.correlationId === null || logEntry.correlationId === '00000000-0000-0000-0000-000000000000')) {
                                    var notificationGroup = result.filter((nle: any) => nle.correlationId == logEntry.correlationId && nle.itemType.toLowerCase() === logEntry.itemType.toLowerCase());
                                    var routeList: string[] = [];

                                    for (const notificationLogEntry of notificationGroup) {
                                        if (notificationLogEntry.extendedProperties.Route)
                                            routeList.push(notificationLogEntry.extendedProperties.Route);
                                        processedIds.push(notificationLogEntry.id);
                                    }
                                    logEntry.extendedProperties.Route = routeList;
                                } else {
                                    processedIds.push(logEntry.id);
                                }

                                items.push({
                                    key: logEntry.id,
                                    id: logEntry.id,
                                    timestamp: new Date(logEntry.timestamp),
                                    user: logEntry.user,
                                    activity: logEntry.activity,
                                    item: logEntry.item,
                                    itemType: logEntry.itemType,
                                    itemUniqueId: logEntry.itemUniqueId,
                                    eventSource: logEntry.eventSource,
                                    extendedProperties: logEntry.extendedProperties,
                                    modifiedProperties: logEntry.modifiedProperties
                                });
                            }                         
                        }

                        items = items.sort((a, b) => 0 - (a.timestamp > b.timestamp ? 1 : -1));
                    }).then(() => {
                        this.setState({
                            items: items,
                            isLoaded: true
                        });
                    });
                }).catch((err) => {
                    this.setState({
                        error: err.message
                    });
                });
    }

    private _getKey(item: any, index?: number): string {
        return item.key;
    } 

    private _onItemInvoked(item: any): void {
        
    }

    private _openViewChangesDialog(item: IAuditLogEntry): void {
        this.setState({
            selectedAuditLogEntry: item
        });
    }

    private _closeViewChangesDialog(parent: any): void {
        parent.setState({
            selectedAuditLogEntry: undefined
        });
    }
}
