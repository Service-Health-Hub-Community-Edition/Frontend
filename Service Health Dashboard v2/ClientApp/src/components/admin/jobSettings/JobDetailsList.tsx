import * as React from 'react';
import {
    Text, FontIcon, Spinner, SpinnerSize, IGroup, DetailsList, SelectionMode, Label, DetailsListLayoutMode, IColumn,
    Dialog, DialogType, DialogFooter, PrimaryButton, Link, ILinkStyles, ILinkStyleProps, mergeStyles, mergeStyleSets
} from '@fluentui/react';
import { MetaDataList, MetaDataItem } from '@m365-admin/metadata';
import type { ITagStyle } from '@m365-admin/tag';
import { Tag, TagType } from '@m365-admin/tag';
import Editor from "@monaco-editor/react";
import { AccessDenied } from "../../AccessDenied";
import { acquireAccessToken } from "../../../auth/AccessTokenHelper";
import { IAuditLogEntry } from '../../AuditLog';

interface IJobDetailsListState {
    id: string;
    items: IAuditLogEntry[];
    groups: IGroup[];
    detailsEntry?: IAuditLogEntry;
    initialized: boolean;
    accessGranted: boolean;
    error?: string;
}

const pipeFabricStyles = (p: ILinkStyleProps): ILinkStyles => ({
    root: {
        textDecoration: 'none',
        color: p.theme.semanticColors.bodyText,
        fontWeight: '600',
        fontSize: p.theme.fonts.medium.fontSize,
    },
});

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

export class JobDetailsList extends React.Component<{ id: string }, IJobDetailsListState> {

    constructor(props: {id: string}) {
        super(props);

        this.state = {
            id: props.id,
            items: [],
            groups: [],
            accessGranted: false,
            initialized: false
        };
    }

    public render() {
        const { items, groups, initialized, detailsEntry, error, accessGranted } = this.state;

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
                <AccessDenied />
            </>);

        const columns: IColumn[] = [
            {
                key: '1',
                name: 'Time',
                fieldName: 'timestamp',
                minWidth: 120,
                isResizable: true,
                isRowHeader: true,
                onRender: (item: IAuditLogEntry) => {
                    return <Link onClick={(event) => {
                        event.preventDefault();
                        this._onOpenDetails(item);
                    }} styles={pipeFabricStyles}>{new Date(item.timestamp).toLocaleString()}</Link >
                },
            },
            {
                key: '2',
                name: 'Type',
                fieldName: 'itemType',
                minWidth: 96,
                maxWidth: 96,
                isResizable: false
            },
            {
                key: '3',
                name: 'Activity',
                isResizable: false,
                minWidth: 128,
                maxWidth: 128,
                onRender: (item: IAuditLogEntry) => {
                    var tagType: TagType = TagType.LowImpactInformational;

                    if (item.activity.toLowerCase().indexOf('failed') > -1)
                        tagType = TagType.ActiveWarning
                    else
                        switch (true) {
                            case (item.activity.toLowerCase().indexOf('job') > -1):
                                tagType = TagType.LowImpactInformational;
                                break;
                            case (item.activity.toLowerCase().indexOf('created') > -1):
                                tagType = TagType.HighImpactInformational;
                                break;
                            case (item.activity.toLowerCase().indexOf('modified') > -1):
                                tagType = TagType.MediumImpactInformational;
                                break;
                            case (item.activity.toLowerCase().indexOf('sent') > -1):
                                tagType = TagType.Recommendation;
                                break;
                            default:
                                tagType = TagType.LowImpactInformational;
                        }
                    
                    return (<Tag key={'tag-' + item.id} tagType={tagType}>{item.activity}</Tag>)
                }
            },
            {
                key: '4',
                name: 'Source',
                fieldName: 'eventSource',
                minWidth: 128,
                maxWidth: 128,
                isResizable: false,
            }
        ];

        return (<>
            <DetailsList
                selectionMode={SelectionMode.none}
                layoutMode={DetailsListLayoutMode.justified}
                compact={true}
                items={items}
                columns={columns}
                groups={groups}
            />
            <Dialog
                title='Details'
                type={DialogType.largeHeader}
                onDismiss={() => this._onCloseDetails()}
                modalProps={{
                    isBlocking: true
                }}
                isOpen={detailsEntry !== undefined}
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
                <MetaDataList>
                    <MetaDataItem header='Item' body={detailsEntry?.itemUniqueId} />
                    <MetaDataItem header='Source' body={detailsEntry?.eventSource} />
                    <MetaDataItem header='Type' body={detailsEntry?.itemType} />
                    <MetaDataItem header='Activity' body={detailsEntry?.activity} />                 
                    <MetaDataItem header='Time' body={detailsEntry?.timestamp ? new Date(detailsEntry?.timestamp).toLocaleString() : ''} />
                </MetaDataList><br />

                {detailsEntry?.extendedProperties ? (<>
                    <Label>Additional information</Label>
                    <Editor
                        height="200px"
                        language="json"
                        value={detailsEntry?.extendedProperties ? JSON.stringify(detailsEntry?.extendedProperties, null, 4) : ''}
                        options={{ automaticLayout: true, readOnly: true, formatOnPaste: true, formatOnType: true }}
                    />
                </>) : ''}

                <DialogFooter>
                    <PrimaryButton onClick={() => this._onCloseDetails()} text="Close" />
                </DialogFooter>

            </Dialog>
        </>);
    }

    componentDidMount() {
        this._loadData();
    }

    _loadData() {
        const requiredRoles: string[] = ['Admin'];
        var items: IAuditLogEntry[] = [];
        var groups: IGroup[] = [];

        this.setState({
            initialized: false
        });   

        acquireAccessToken()
            .then((response) => {
                var tokenClaims: any = response.account?.idTokenClaims;
                const userRoles: any = tokenClaims?.roles;
                var userHasRequiredRole: boolean = userRoles.some((r: string) => requiredRoles.includes(r));

                this.setState({
                    accessGranted: userHasRequiredRole
                });

                if (userHasRequiredRole)
                    fetch('/api/AuditLog?correlationId=' + this.state.id, { headers: response.idToken === "" ? {} : { 'Authorization': `Bearer ${response.idToken}` } })
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

                            items = items.sort((a, b) => 0 - (a.timestamp > b.timestamp ? 1 : -1));
                            items = items.sort((a, b) => (a.itemUniqueId > b.itemUniqueId ? 1 : -1));
                            const groupNames: string[] = Array.from(new Set(items.map(i => i.itemUniqueId)));
                            groups = this._getGroups(items, groupNames);

                        }).then(() => {
                            this.setState({
                                items: items,
                                groups: groups,
                                initialized: true
                            });
                        });
            }).catch((err) => {
                this.setState({
                    error: err.message,
                    initialized: true
                });
            });
    }

    private _getGroups(groupItems: IAuditLogEntry[], groupNames: string[]): IGroup[] {
        const serviceGroups: IGroup[] = [];
        let groupsStartIndex = 0;

        groupNames.forEach((itemId: string) => {
            const endpointSets = groupItems.filter((set: IAuditLogEntry) => {
                return set.itemUniqueId.toLowerCase() === itemId.toLowerCase();
            });

            const setGroupCount = endpointSets.length;

            const group: IGroup = {
                key: itemId,
                name: itemId,
                startIndex: groupsStartIndex,
                count: setGroupCount
            };

            groupsStartIndex += setGroupCount;

            serviceGroups.push(group);
        });

        return serviceGroups;
    }

    _onOpenDetails(item: IAuditLogEntry): void {
        this.setState({
            detailsEntry: item
        });
    }

    _onCloseDetails(): void {
        this.setState({
            detailsEntry: undefined
        });
    }
}