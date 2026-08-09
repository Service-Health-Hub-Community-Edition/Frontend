import * as React from 'react';
import {
    Text, Spinner, SpinnerSize, IGroup, DetailsList, SelectionMode, Label, DetailsListLayoutMode, IColumn,
    Dialog, DialogType, DialogFooter, PrimaryButton, Link, ILinkStyles, ILinkStyleProps, mergeStyles, mergeStyleSets, IContextualMenuProps, IconButton,
    DefaultButton, Selection, IObjectWithKey, MessageBar, MessageBarType, Icon, CommandBar
} from '@fluentui/react';
import { MetaDataList, MetaDataItem } from '@m365-admin/metadata';
import type { ITagStyle } from '@m365-admin/tag';
import { Tag, TagType } from '@m365-admin/tag';
import { AdminLayout } from '../AdminLayout';
import { DetailPageHeader } from '@m365-admin/detail-page';
import { M365Breadcrumb } from '@m365-admin/m365-breadcrumb';
import { AccessDenied } from "../../AccessDenied";
import { acquireAccessToken } from "../../../auth/AccessTokenHelper";
import { TextField } from 'office-ui-fabric-react';

export interface ITagDefinition {
    key: string;
    id: number;
    tagId: string;
    name: string;
    type: string;
    itemCount: number;
    lastUsed: Date | undefined | null;
}

interface ITagDefinitionsState {
    items: ITagDefinition[];
    groups: IGroup[];
    renameEntry?: string | undefined;
    newName?: string | undefined;
    detailsEntry?: ITagDefinition;
    moveDialogOpen: boolean;
    targetTagGroup?: string | number | undefined;
    initialized: boolean;
    accessGranted: boolean;
    error?: string;
    saveError?: string | undefined;
    updateInProgress: boolean;
}

const pipeFabricStyles = (p: ILinkStyleProps): ILinkStyles => ({
    root: {
        textDecoration: 'none',
        color: p.theme.semanticColors.bodyText,
        fontWeight: 'normal',
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

export class TagDefinitionsList extends React.Component<{}, ITagDefinitionsState> {
    private _targetGroupSelection: Selection;
    private _groupNames: Map<string, string> = new Map<string, string>();

    constructor(props: {}) {
        super(props);

        this._targetGroupSelection = new Selection({
            onSelectionChanged: () => this.setState({ targetTagGroup: this._getTargetGroupSelectionDetails() }),
        });

        this.state = {
            items: [],
            groups: [],
            moveDialogOpen: false,
            targetTagGroup: undefined,
            accessGranted: false,
            initialized: false,
            updateInProgress: false
        };
    }

    public render() {
        const { items, groups, initialized, detailsEntry, error, accessGranted, moveDialogOpen, targetTagGroup, newName, renameEntry, updateInProgress, saveError } = this.state;
        
        if (accessGranted === false)
            return (<>
                <AccessDenied />
            </>);

        const columns: IColumn[] = [
            {
                key: '1',
                name: 'Name',
                fieldName: 'name',
                minWidth: 120,
                isResizable: true,
                onRender: (item: ITagDefinition) => {
                    if (item.tagId.toLowerCase() === renameEntry?.toLowerCase() && newName === undefined)
                        this._setNewName(item.name);

                    return item.tagId.toLowerCase() === renameEntry?.toLowerCase() ? (<>
                        <TextField
                            defaultValue={item.name}
                            onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string | undefined) => {
                                this._setNewName(newValue)
                            }}
                            minLength={2}
                            maxLength={64}
                            required
                        />
                        <IconButton
                            iconProps={{ iconName: 'Accept', className: classNames.green }}
                            disabled={newName ? newName.length < 2 : true}
                            onClick={() => { this._addRenameTag(item, newName) }}
                        />
                        <IconButton
                            iconProps={{ iconName: 'Cancel', className: classNames.red }}
                            onClick={() => { this._onCancelRename(item); }}
                        />
                        {updateInProgress ? (<Spinner size={SpinnerSize.small} />) : (<></>)}
                        {saveError ? (<MessageBar
                            messageBarType={MessageBarType.error}
                            isMultiline={false}
                            dismissButtonAriaLabel="Close"
                        >
                            Couldn't update tag: {saveError}
                        </MessageBar>) : (<></>)}
                    </>
                    ) : (<><Icon iconName='Tag' /> <Text variant='medium'>{item.name}</Text></>)
                },
            },
            {
                key: 'clSHHItemMenu',
                name: '',
                minWidth: 8,
                maxWidth: 8,
                isResizable: false,
                isCollapsible: true,
                isPadded: false,
                styles: {

                    root: { padding: '6px 0px 6px 0px' }
                },
                onRender: (item: ITagDefinition) => {
                    var menuProps: IContextualMenuProps = {
                        directionalHintFixed: true,
                        items: [
                            {
                                key: item.tagId + '-Rename',
                                text: 'Rename',
                                iconProps: { iconName: 'Edit' },
                                onClick: () => { this._onRenameTag(item) }
                            },
                            {
                                key: item.tagId + '-Move',
                                text: 'Move',
                                iconProps: { iconName: 'MoveToFolder' },
                                onClick: () => { this._onOpenMoveTag(item) }
                            },
                            {
                                key: item.tagId + '-Delete',
                                text: 'Delete',
                                iconProps: { iconName: 'Delete' },
                                onClick: () => { this._onDeleteTag(item) }
                            }
                        ]
                    };

                    return (<>
                        <IconButton
                            style={{ height: '16px' }}
                            iconProps={{ iconName: 'MoreVertical' }}
                            menuProps={menuProps}
                            menuIconProps={{ hidden: true }} />
                    </>)
                }
            },
            {
                key: '2',
                name: 'Item count',
                fieldName: 'itemCount',
                minWidth: 96,
                maxWidth: 96,
                isResizable: false
            },
            {
                key: '3',
                name: 'Last used',
                isResizable: false,
                minWidth: 128,
                maxWidth: 128,
                onRender: (item: ITagDefinition) => {
                    return (<>{item.lastUsed !== null && item.lastUsed !== undefined ? new Date(item.lastUsed).toLocaleString() : ""}</>)
                }
            }
        ];

        return (<AdminLayout>
            <div className="container" style={{ maxWidth: '97%' }}>
                <div className="row">
                    <div className="col">
                        <M365Breadcrumb
                            items={[
                                { text: 'Home', key: 'home', href: '/admin' },
                                { text: 'Tags', key: 'tags', href: '/admin/tags', isCurrentItem: true },
                            ]}
                            style={{ marginBottom: '16px' }}
                        />
                        <DetailPageHeader
                            title="Tag definitions"
                            description="Create, edit or remove user-defined tags. Tagging communication items helps you quickly filter the service communications by categories that you define."
                        />
                    </div>
                </div>
                {!initialized ? (<>
                    <br /><br />
                    <Spinner size={SpinnerSize.medium} />
                </>) :
                    error ? (<>
                <br /><br />
                <Text variant='medium'><b>{error}</b></Text>
            </>) : (
                <div className="row">
                    <div className="col">
                        <div>
                            <CommandBar items={[
                                {
                                    key: 'newTag',
                                    text: 'Add tag',
                                    cacheKey: 'myCacheKey', // changing this key will invalidate this item's cache
                                    iconProps: { iconName: 'Add' },
                                    subMenuProps: {
                                        items: groups.map((g: IGroup) => {
                                            return {
                                                key: g.key,
                                                text: g.name,
                                                onClick: () => {this._addTagPlaceholder(g.key)}
                                                }
                                            })
                                    }
                                }
                            ]} /> 

                            <DetailsList
                                key='lst_TagDefs'
                                selectionMode={SelectionMode.none}
                                layoutMode={DetailsListLayoutMode.justified}
                                compact={true}
                                items={items}
                                columns={columns}
                                groups={groups}
                                groupProps={{ showEmptyGroups: true }}
                                onShouldVirtualize={() => false}
                            />
                        </div>
                    </div>
                </div>)}
            </div>
            <Dialog
                title={'Move tag' + (detailsEntry ? ' - ' + detailsEntry.name : '')}
                type={DialogType.largeHeader}
                onDismiss={() => this._onCloseMoveTag()}
                modalProps={{
                    isBlocking: true
                }}
                isOpen={moveDialogOpen === true}
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
                onLayerDidMount={() => {
                    if (detailsEntry !== undefined) {
                        this._targetGroupSelection.setAllSelected(false);
                        this._targetGroupSelection.setKeySelected(detailsEntry.type.toLowerCase(), true, false);
                    }
                }}
            >
                <DetailsList
                    key='lst_TagGroups'
                    selectionMode={SelectionMode.single}
                    layoutMode={DetailsListLayoutMode.justified}
                    compact={true}
                    items={groups}
                    columns={[
                        {
                            key: 'tg_1',
                            name: 'Target group',
                            minWidth: 120,
                            isResizable: true,
                            onRender: (item: IGroup) => {
                                return item.name
                            },
                        }]
                    }
                    groupProps={{ showEmptyGroups: true }}
                    onShouldVirtualize={() => false}
                    selection={this._targetGroupSelection}
                />

                <DialogFooter>
                    {updateInProgress ? (<Spinner size={SpinnerSize.small} />) : (<></>)}
                    {saveError ? (<><MessageBar
                        messageBarType={MessageBarType.error}
                        isMultiline={true}
                        dismissButtonAriaLabel="Close"
                    >
                        Couldn't update tag: {saveError}
                    </MessageBar><br/></>) : (<></>)}
                    <PrimaryButton onClick={() => this._onMoveTag()} disabled={targetTagGroup === undefined || detailsEntry?.type.toLowerCase() === targetTagGroup?.toString().toLowerCase()} text="Move" />
                    <DefaultButton onClick={() => this._onCloseMoveTag()} text="Cancel" />
                </DialogFooter>

            </Dialog>
        </AdminLayout>
        );
    }

    componentDidMount() {
        this._loadData();
    }

    _loadData() {
        const requiredRoles: string[] = ['Admin'];
        var items: ITagDefinition[] = [];
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
                    fetch('/api/TagDefinitions', { headers: response.idToken === "" ? {} : { 'Authorization': `Bearer ${response.idToken}` } })
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
                            for (const groupedTagDefinition of result) {
                                if (groupedTagDefinition.tagDefinitions !== null && groupedTagDefinition.tagDefinitions !== undefined) {
                                    for (const tagDefinition of groupedTagDefinition.tagDefinitions)
                                        items.push({
                                            key: tagDefinition.tagId,
                                            id: tagDefinition.id,
                                            name: tagDefinition.name,
                                            tagId: tagDefinition.tagId,
                                            type: tagDefinition.type && tagDefinition.type !== null ? tagDefinition.type : "!!!general",
                                            lastUsed: tagDefinition.lastUsed,
                                            itemCount: tagDefinition.itemCount
                                        });
                                }

                                if (groupedTagDefinition.internalName === null) {
                                    this._groupNames.set("!!!general", groupedTagDefinition.component);
                                } else {
                                    this._groupNames.set(groupedTagDefinition.internalName.toLowerCase(), groupedTagDefinition.component);
                                }
                            }

                            items = items.sort((a, b) => a.name > b.name ? -1 : 1);
                            items = items.sort((a, b) => a.type > b.type ? 1 : -1);
                            groups = this._getGroups(items, this._groupNames);

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

    private _getGroups(groupItems: ITagDefinition[], groups: Map<string, string>): IGroup[] {
        const serviceGroups: IGroup[] = [];
        let groupsStartIndex = 0;
        const groupNames: Map<string, string> = new Map(Array.from(groups).sort());

        groupNames.forEach((value: string, key: string) => {
            const tagDefinitions = groupItems.filter((set: ITagDefinition) => {
                return set.type?.toLowerCase() === key?.toLowerCase();
            });

            const setGroupCount = tagDefinitions.length;

            const group: IGroup = {
                key: key,
                name: value,
                startIndex: groupsStartIndex,
                count: setGroupCount
            };

            groupsStartIndex += setGroupCount;

            serviceGroups.push(group);
        });

        return serviceGroups;
    }

    _onOpenDetails(item: ITagDefinition): void {
        this.setState({
            detailsEntry: item
        });
    }

    _onCloseDetails(): void {
        this.setState({
            detailsEntry: undefined
        });
    }

    _onOpenMoveTag(item: ITagDefinition): void {
        this.setState({
            targetTagGroup: item.type.toLowerCase(),
            detailsEntry: item,
            moveDialogOpen: true
        });

        this._targetGroupSelection.setAllSelected(false);
        this._targetGroupSelection.setKeySelected(item.type.toLowerCase(), true, false);
    }


    _getTargetGroupSelectionDetails(): string | number | undefined {
        const selection: IObjectWithKey[] = this._targetGroupSelection.getSelection();
        const result: string | number | undefined = selection.length > 0 ? selection[0].key : undefined;
        console.log(result);
        return result;
    }

    _addTagPlaceholder(groupName: string): void {
        var items: ITagDefinition[] = this.state.items;
        var item: ITagDefinition = {
            id: -1,
            tagId: '00000000-0000-0000-0000-000000000000',
            itemCount: 0,
            key: '00000000-0000-0000-0000-000000000000',
            lastUsed: null,
            name: '',
            type: groupName
        };

        items.push(item);

        items = items.sort((a, b) => a.name > b.name ? -1 : 1);
        items = items.sort((a, b) => a.type > b.type ? 1 : -1);
        var groups: IGroup[] = this._getGroups(items, this._groupNames);

        this.setState({
            newName: '',
            renameEntry: item.tagId,
            saveError: undefined,
            items: items,
            groups: groups
        });
    }

    _moveTag(item: ITagDefinition | undefined, targetGroup: string | number | undefined): void {
        if (item === undefined || targetGroup === undefined)
            return;

        var params: RequestInit;

        this.setState({
            updateInProgress: true
        });

        var tagOpBody = {
            op: 'move',
            data: { targetGroup: targetGroup.toString() === '!!!general' ? null : targetGroup.toString() }
        };

        acquireAccessToken()
            .then((response) => {
                params = {
                    headers: {
                        "Content-Type": "application/json charset=UTF-8",
                        "Authorization": "Bearer " + response.idToken
                    },
                    body: JSON.stringify(tagOpBody),
                    method: "POST"
                };

                var r: Response;

                fetch("/api/TagDefinitions/" + item.tagId, params)
                    .then((response: Response) => {
                        r = response;
                        return response.text();
                    }).then((data) => {
                        if (r.ok) {
                            this.setState({
                                targetTagGroup: undefined,
                                moveDialogOpen: false,
                                updateInProgress: false
                            });

                            this._loadData();
                        } else {
                            this.setState({
                                updateInProgress: false,
                                saveError: r.status + ': ' + data
                            });
                        }
                    })
                    .catch((err) => {
                        this.setState({
                            saveError: err.message,
                            updateInProgress: false
                        });
                    });
            })
            .catch((err) => {
                this.setState({
                    saveError: "Authentication error. Details:<br/><br/>" + err.message,
                    updateInProgress: false
                });
            });     
    }

    _onMoveTag(): void {
        if (this.state.targetTagGroup !== undefined) {
            if (this.state.targetTagGroup !== "!!!general") {
                if (window.confirm('Moving tag to a different group will remove the tag from all currently tagged items. Proceed with caution.')) {
                    this._moveTag(this.state.detailsEntry, this.state.targetTagGroup);
                } else {
                    return;
                }
            } else {
                this._moveTag(this.state.detailsEntry, this.state.targetTagGroup);
            }
        }
    }

    _onCloseMoveTag(): void {
        this._targetGroupSelection.setAllSelected(false);
        this.setState({
            targetTagGroup: undefined,
            moveDialogOpen: false,
            saveError: undefined
        });
    }

    _onRenameTag(item: ITagDefinition): void {
        this.setState({
            renameEntry: item.tagId
        });
    }

    _setNewName(newName?: string | undefined): void {
        this.setState({
            newName: newName ? newName : ''
        });
    }

    _addRenameTag(item: ITagDefinition, newName: string | undefined): void {
        if (newName && newName.trim() !== '' && newName.trim() !== item.name.trim()) {

            var params: RequestInit;

            this.setState({
                updateInProgress: true
            });

            var tagOpBody = item.id === -1 ? {
                    op: 'create',
                    data: {
                        name: newName.trim(),
                        targetGroup: item.type.toLowerCase() === '!!!general' ? null : item.type
                    }
                } :
                {
                    op: 'rename',
                    data: { name: newName.trim() }
                }

            acquireAccessToken()
                .then((response) => {
                    params = {
                        headers: {
                            "Content-Type": "application/json charset=UTF-8",
                            "Authorization": "Bearer " + response.idToken
                        },
                        body: JSON.stringify(tagOpBody),
                        method: "POST"
                    };

                    var r: Response;

                    var uri: string = "/api/TagDefinitions";
                    if (item.id !== -1)
                        uri += "/" + item.tagId

                    fetch(uri, params)
                        .then((response: Response) => {
                            r = response;
                            return response.text();
                        }).then((data) => {
                            if (r.ok) {
                                this.setState({
                                    renameEntry: undefined,
                                    newName: undefined,
                                    updateInProgress: false
                                });

                                this._loadData();
                            } else {
                                this.setState({
                                    updateInProgress: false,
                                    saveError: r.status + ': ' + data
                                });
                            }
                        })
                        .catch((err) => {
                            this.setState({
                                saveError: err.message,
                                updateInProgress: false
                            });
                        });
                })
                .catch((err) => {
                    this.setState({
                        saveError: "Authentication error. Details:<br/><br/>" + err.message,
                        updateInProgress: false
                    });
                }); 
        }
    }

    _onCancelRename(item: ITagDefinition): void {
        if (item.id === -1) {
            var items: ITagDefinition[] = this.state.items.filter((i: ITagDefinition) => i.id !== -1);

            items = items.sort((a, b) => a.name > b.name ? -1 : 1);
            items = items.sort((a, b) => a.type > b.type ? 1 : -1);
            var groups: IGroup[] = this._getGroups(items, this._groupNames);

            this.setState({
                newName: undefined,
                renameEntry: undefined,
                saveError: undefined,
                items: items,
                groups: groups
            });

        } else {
            this.setState({
                newName: undefined,
                renameEntry: undefined,
                saveError: undefined
            });
        }
    }

    _onDeleteTag(item: ITagDefinition): void {
        if (window.confirm('Deleting tag will also remove the tag from all currently tagged items. This operation cannot be undone. Proceed with caution.')) {
            this._removeTag(item);
        } else {
            return;
        }
    }

    _removeTag(item: ITagDefinition): void {
        var params: RequestInit;

        this.setState({
            updateInProgress: true
        });

        acquireAccessToken()
            .then((response) => {
                params = {
                    headers: {
                        "Content-Type": "application/json charset=UTF-8",
                        "Authorization": "Bearer " + response.idToken
                    },
                    method: "DELETE"
                };

                var r: Response;

                fetch("/api/TagDefinitions/" + item.tagId, params)
                    .then((response: Response) => {
                        r = response;
                        return response.text();
                    }).then((data) => {
                        if (r.ok) {
                            this._loadData();
                        } else {
                            this.setState({
                                updateInProgress: false,
                                saveError: r.status + ': ' + data
                            });
                        }
                    })
                    .catch((err) => {
                        this.setState({
                            saveError: err.message,
                            updateInProgress: false
                        });
                    });
            })
            .catch((err) => {
                this.setState({
                    saveError: "Authentication error. Details:<br/><br/>" + err.message,
                    updateInProgress: false
                });
            });
    }
}