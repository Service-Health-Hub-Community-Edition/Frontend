import * as React from 'react';
import {
    CommandBar, ICommandBarItemProps, IContextualMenuItem, Text, Icon, Pivot, PivotItem, FontIcon,
    DefaultPalette, mergeStyles, mergeStyleSets, ITag, TagPicker, IBasePickerSuggestionsProps
} from '@fluentui/react';
import { Separator } from '@fluentui/react';
import { Dropdown, IDropdownOption, IDropdownStyles, Button } from '@fluentui/react';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { MessageBar, MessageBarType, TextField, Dialog, DialogType, DialogFooter, PrimaryButton, TooltipHost } from '@fluentui/react';
import { MetaDataList, MetaDataItem } from '@m365-admin/metadata';
import { ICustomAction, CustomAction } from './CustomAction';
import { AuditLog } from './AuditLog';
import { ServiceComponent } from './ServiceNameComponent';
import { AlertMessage } from './AlertMessage';
import { acquireAccessToken } from "../auth/AccessTokenHelper";
import { setViewState, setFavoriteState, setArchiveState } from '../api/viewpoint';
import { getAvailableTagDefinitions, getTags, addTag, removeTag, MSTag } from '../api/tags';
import { ITagDefinition } from './admin/applicationSettings/TagDefinitions';

export interface IIncidentDetailsState {
    incidentId: string;
    incidentTitle: string;
    service: string[];
    status: string;
    issueType?: string;
    origin?: string;
    userImpact: string;
    latestMessage: string;
    allMessages: ServiceMessage[];
    messagesDropdownOptions: IDropdownOption[],
    selectedMessageItem?: IDropdownOption;
    lastUpdated: Date;
    startTime: Date;
    endTime?: Date;
    workItem: string;
    publicComm: boolean;
    summary?: any; 
    comments: string;
    newComments: string;
    serviceHealthHubViewpoint: any;
    customActions: ICustomAction[] | undefined;
    canWriteMetadata: boolean;
    selectedTags: ITag[];
    availableTags: ITag[];
    tagError?: string;
    isPublishMessageDialogOpen: boolean;
    loading: boolean;
    publishingError?: string;
    error?: string;
}

interface ServiceMessageData {
    content: string;
}

interface ServiceMessage {
    description: ServiceMessageData;
    createdDateTime: string;
}

const dropdownStyles: Partial<IDropdownStyles> = { dropdown: { width: 200 } };
const iconClass = mergeStyles({

});

const classNames = mergeStyleSets({
    incident: [{ color: 'red' }, iconClass],
    advisory: [{ color: DefaultPalette.blue }, iconClass]
});

const componentName: string = 'ServiceHealthIssue';

export class IncidentDetails extends React.Component<{
    id?: string,
    onPublishingChange?: any, 
    onView?: any,
    onFavorite?: any,
    onArchive?: any,
    onUpdateOrgTags?: any
    onUpdateParentTitle?: any
}, IIncidentDetailsState> {

    customAction: any = React.createRef();
    activityLog: any = React.createRef();

    constructor(props: {
        id?: string,
        onPublishingChange?: any,
        onView?: any,
        onFavorite?: any,
        onArchive?: any,
        onUpdateOrgTags?: any,
        onUpdateParentTitle?: any
    }) {
        super(props);

        this.state = {
            incidentId: "",
            incidentTitle: "",
            service: [],
            status: "",
            userImpact: "",
            latestMessage: "",
            allMessages: [],
            messagesDropdownOptions: [],
            selectedMessageItem: undefined,
            lastUpdated: new Date(),
            startTime: new Date(),
            endTime: undefined,
            workItem: "",
            publicComm: false,
            comments: "",
            newComments: "",
            serviceHealthHubViewpoint: undefined,
            customActions: undefined,
            canWriteMetadata: false,
            selectedTags: [],
            availableTags: [],
            isPublishMessageDialogOpen: false,
            loading: false,
            publishingError: undefined,
            error: undefined
        };
    }

    public render() {
        const {
            incidentId, service, status, userImpact, latestMessage,
            lastUpdated, startTime, workItem, allMessages, messagesDropdownOptions, issueType, origin,
            selectedMessageItem, publicComm, newComments, summary, serviceHealthHubViewpoint, customActions,
            isPublishMessageDialogOpen, loading, publishingError, error,
            canWriteMetadata, selectedTags, availableTags, tagError } = this.state;

        var _items: ICommandBarItemProps[] = [];
        var customActionItems: IContextualMenuItem[] = [];

        if (customActions)
            for (const customAction of customActions) {
                customAction.icon.trim() !== "" ?
                    customActionItems.push({
                        key: customAction.actionId,
                        text: customAction.name,
                        iconProps: { iconName: customAction.icon.trim() },
                        onClick: () => this.customAction.current._onRunCustomAction(customAction.name, customAction.actionId, incidentId),
                    }) :
                    customActionItems.push({
                        key: customAction.actionId,
                        text: customAction.name,
                        onClick: () => this.customAction.current._onRunCustomAction(customAction.name, customAction.actionId, incidentId),
                    })
            }

        // add publishing commands
        const _publishItems: ICommandBarItemProps[] = [
            {
                key: 'publish',
                text: 'Publish',
                iconProps: { iconName: 'PublishContent' },
                onClick: () => this._onOpenPublishMessageDialog(),
            },
            {
                key: 'update',
                text: 'Update',
                iconProps: { iconName: 'PostUpdate' },
                disabled: true
            }
        ];

        const _updateItems: ICommandBarItemProps[] = [
            {
                key: 'unpublish',
                text: 'Unpublish',
                iconProps: { iconName: 'UnpublishContent' },
                onClick: () => this._procecssUnpublishing(),
            },
            {
                key: 'update',
                text: 'Update',
                iconProps: { iconName: 'PostUpdate' },
                onClick: () => this._onOpenPublishMessageDialog(),
            }           
        ];

        const publishingActions: ICommandBarItemProps[] = publicComm ? _updateItems : _publishItems;

        for (const publishingAction of publishingActions)
            _items.push(publishingAction)

        // add custom actions
        if (customActionItems.length > 0)
            _items.push({
                key: "itemIntegrate",
                text: "Integrate",
                iconProps: { iconName: "AppIconDefault" },
                subMenuProps: { items: customActionItems }
            });  

        if (serviceHealthHubViewpoint) {
            _items.push(serviceHealthHubViewpoint.viewed ?
                {
                    key: 'markAsUnread',
                    text: 'Mark as unread',
                    iconProps: { iconName: 'Mail' },
                    onClick: () => this._setViewState(incidentId, false)
                } : {
                    key: 'markAsRead',
                    text: 'Mark as read',
                    iconProps: { iconName: 'Read' },
                    onClick: () => this._setViewState(incidentId, true)
                });

            _items.push(serviceHealthHubViewpoint.favorite ?
                {
                    key: 'removeFavorite',
                    text: 'Remove from favorites',
                    iconProps: { iconName: 'FavoriteStarFill' },
                    onClick: () => this._setFavoriteState(incidentId, false)
                } : {
                    key: 'addFavorite',
                    text: 'Add to favorites',
                    iconProps: { iconName: 'FavoriteStar' },
                    onClick: () => this._setFavoriteState(incidentId, true)
                });
        }

        const getTextFromItem = (item: ITag) => item.name;

        const listContainsTagList = (tag: ITag, tagList?: ITag[]) => {
            if (!tagList || !tagList.length || tagList.length === 0) {
                return false;
            }
            return tagList.some(compareTag => compareTag.key === tag.key);
        };

        const filterSuggestedTags = (filterText: string, tagList: ITag[] | undefined): ITag[] => {
            return filterText
                ? availableTags.filter(
                    tag => tag.name.toLowerCase().indexOf(filterText.toLowerCase()) === 0 && !listContainsTagList(tag, tagList),
                )
                : [];
        };

        const pickerSuggestionsProps: IBasePickerSuggestionsProps = {
            suggestionsHeaderText: 'Suggested organization tags',
            noResultsFoundText: 'No tags found',
        };

        return (
            <div>
                <CustomAction componentName={componentName} ref={this.customAction} onLoad={(actions: ICustomAction[]) => this._onLoadCustomActions(actions)} />

                {loading ? (
                    <div className="loadingProgress">
                        <br /><br />
                        <Spinner size={SpinnerSize.large} />
                    </div>
                ) :
                (
                        <div>
                            <div className="incidentDetails" style={{ display: incidentId ? 'block' : 'none' }}>
                                <div className="container" style={{ padding: '0px' }}>
                                    <div className="row">
                                        <div className="col">
                                            <Text variant={'medium'} >
                                                {incidentId}, Last updated: {lastUpdated.toLocaleDateString([], { year: 'numeric', month: 'short', day: 'numeric', hour: 'numeric', minute: 'numeric' })} <br />
                                                Estimated start time: {startTime.toLocaleDateString([], { year: 'numeric', month: 'short', day: 'numeric', hour: 'numeric', minute: 'numeric' })}
                                            </Text>
                                        </div>
                                    </div>
                                    <div className="row">
                                        <div className="col">
                                            <CommandBar
                                                items={_items}
                                                farItems={[]}
                                                ariaLabel="Service Health Incident actions"
                                            />
                                        </div>
                                    </div>
                                    <div className="row" style={{ paddingTop: "10px", marginBottom: "20px" }}>
                                        <div className="col">
                                            <MetaDataList>
                                                <MetaDataItem header='Service' body={ service !== undefined ? service.map((s) => (<ServiceComponent name={s} />)) : ""} />
                                                <MetaDataItem header='Task' body={workItem !== "" ? (<div><FontIcon iconName='TaskSolid' />&nbsp;<Text variant={'small'} >
                                                    <div style={{ whiteSpace: "pre-wrap", display: 'inline-block' }} dangerouslySetInnerHTML={{ __html: workItem }} />
                                                </Text></div>) : (<Text variant={'small'} >not present</Text>)} />
                                                <MetaDataItem header='Status' body={status} />
                                                <MetaDataItem header='Origin' body={origin} />
                                                <MetaDataItem header='Type' body={
                                                    <div className="container" style={{ margin: "0px" }} >
                                                        <div className="row" style={{ paddingBottom: "0px" }} >
                                                            <div className="col-auto" style={{ margin: "0px", padding: "0px 6px 0px 0px" }} >
                                                                {issueType ? (<FontIcon
                                                                    iconName={issueType === 'Incident' ? 'WarningSolid' : 'InfoSolid'}
                                                                    className={issueType === 'Incident' ? classNames.incident : classNames.advisory} />) : ""}
                                                            </div>
                                                            <div className="col" style={{ margin: "0px", padding: "0px", alignItems: 'center' }}>
                                                                <Text variant='medium'>{issueType}</Text>
                                                            </div>
                                                        </div>
                                                    </div>
                                                } />
                                                <MetaDataItem header='User impact' body={userImpact} />
                                                {availableTags && availableTags.length > 0 ? (
                                                    <MetaDataItem header='Organization tags' body={<>
                                                        <TagPicker
                                                            removeButtonAriaLabel="Remove"
                                                            selectionAriaLabel="Selected tags"
                                                            onResolveSuggestions={filterSuggestedTags}
                                                            getTextFromItem={getTextFromItem}
                                                            pickerSuggestionsProps={pickerSuggestionsProps}
                                                            // this option tells the picker's callout to render inline instead of in a new layer
                                                            pickerCalloutProps={{ doNotLayer: true }}
                                                            onChange={(items?: ITag[] | undefined): void => this._onChangeOrgTags(items)}
                                                            defaultSelectedItems={selectedTags}
                                                            disabled={!canWriteMetadata}
                                                        />
                                                        {tagError ? (<MessageBar
                                                            messageBarType={MessageBarType.error}
                                                            isMultiline={true}
                                                            dismissButtonAriaLabel="Close"
                                                        >
                                                            Couldn't update tag: {tagError}
                                                        </MessageBar>) : (<></>)}
                                                    </>} />) : (<></>)}
                                            </MetaDataList>
                                        </div>
                                    </div>
                                </div>

                                <Pivot
                                    aria-label="Message"
                                    linkFormat={'links'}
                                    overflowBehavior={'menu'}
                                    overflowAriaLabel="more items">

                                    <PivotItem headerText="Details">
                                        &nbsp;<br />
                                        {summary?.contents && summary?.contents.length > 0 ? (
                                            <>
                                                <Text variant='medium'><b>Summary</b></Text>
                                                <TooltipHost content='Machine-generated communication summary'>
                                                    <Icon iconName='Info' style={{paddingLeft: '6px', cursor: 'default', fontSize: 'smaller'}} />
                                                </TooltipHost>
                                                <br />
                                                {summary.timestamp ? (
                                                    <Text variant='xSmall'>Updated: {new Date(summary.timestamp).toLocaleString()}</Text>) : ""
                                                }
                                                {summary?.contents.length > 1 ? (
                                                    <ul style={{ paddingTop: "6px" }} >
                                                        {summary.contents.map((summaryLine: any) => (
                                                            <li style={{ paddingTop: "6px" }}><Text variant='medium'>{summaryLine.text}</Text></li>
                                                        ))}
                                                    </ul>) : (<div><br />
                                                        <Text variant='medium'>{summary.contents[0].text}</Text></div>
                                                )}
                                            </>
                                        ) : "" }

                                        <Text variant={'medium'} block>
                                            <br /><b>Latest message</b>
                                        </Text>
                                        <Text>
                                            <div style={{ whiteSpace: "pre-wrap" }} dangerouslySetInnerHTML={{ __html: latestMessage }} />
                                        </Text>
                                        <div className="incidentMessages" style={{ display: this.state.allMessages.length > 0 ? 'block' : 'none' }}>
                                            <Separator />
                                            <Dropdown
                                                label="Message history"
                                                selectedKey={selectedMessageItem ? selectedMessageItem.key : undefined}
                                                // eslint-disable-next-line react/jsx-no-bind
                                                onChange={this._onMessageDropdownChange}
                                                placeholder="Select a message"
                                                options={messagesDropdownOptions}
                                                styles={dropdownStyles}
                                                isDisabled={messagesDropdownOptions.length === 0}
                                            />
                                            <Text>
                                                <div style={{ whiteSpace: "pre-wrap" }} dangerouslySetInnerHTML={{ __html: selectedMessageItem ? allMessages.find(m => m.createdDateTime === selectedMessageItem.key.toString())?.description.content! : "" }} />
                                            </Text>
                                        </div>
                                        <Separator />
                                    </PivotItem>

                                    <PivotItem headerText="Activity">
                                        &nbsp;<br />
                                        {incidentId ? (<AuditLog itemId={incidentId} itemType={componentName} scheme='item' />) : (<Text variant='medium'>No activities available</Text>)}
                                    </PivotItem>
                                </Pivot>
                            </div>

                            <AlertMessage title='' header='Something went wrong.' message={error} isBlocking={true} isOpen={error !== undefined} onClose={this._closeErrorDialog} />
                            <AlertMessage title='' header='Something went wrong.' message={publishingError} isBlocking={true} isOpen={publishingError !== undefined} onClose={this._closePublishingErrorDialog} />

                            <Dialog
                                hidden={!isPublishMessageDialogOpen}
                                onDismiss={this._onDismisPublishMessageDialog}
                                dialogContentProps={{
                                    type: DialogType.normal,
                                    title: 'Publish incident',
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
                                <TextField
                                    label="Comments"
                                    defaultValue={newComments}
                                    multiline rows={4}
                                    onChange={this._onPublicCommentsChange} />

                                <DialogFooter>
                                    <PrimaryButton onClick={this._procecssPublishing} text="Save" />&nbsp;
                                    <Button onClick={this._cancelPublish} text="Cancel" />
                                </DialogFooter>
                            </Dialog>
                        </div >
                    )}
                </div>);
    }

    componentDidMount() {
        this._onLoadTags(this.props.id, componentName);
        this._getIncidentDetails(this.props.id);
    }

    componentDidUpdate(prevProps: { id?: string }) {
        if (prevProps.id !== this.props.id) {
            this._onLoadTags(this.props.id, componentName);
            this._getIncidentDetails(this.props.id);
        }
    }

    private _onMessageDropdownChange = (event: React.FormEvent<HTMLDivElement>, item?: IDropdownOption): void => {
        this.setState({
            selectedMessageItem: item
        });
    };

    private _getIncidentDetails(id?: string) {
        var incidentId = id;
        if (incidentId !== undefined) {
            this.setState({
                loading: true
            });

            acquireAccessToken()
                .then((response) => {
                    var tokenClaims: any = response.account?.idTokenClaims;
                    const userRoles: any = tokenClaims?.roles;
                    // var userHasRequiredRole: boolean = userRoles.some((r: string) => requiredRoles.includes(r));
                    const commsMgrRoles: string[] = ['Communication.Write.All', 'Admin'];
                    var canWriteMetadata: boolean = userRoles.some((r: string) => commsMgrRoles.includes(r));

                    fetch('/api/issues/' + incidentId, { headers: response.idToken === "" ? {} : { 'Authorization': `Bearer ${response.idToken}` } })
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
                        var sortedMessages = result.posts.sort((a: ServiceMessage, b: ServiceMessage) => 0 - (new Date(a.createdDateTime) > new Date(b.createdDateTime) ? 1 : -1));

                        const incidentMessagesDropdownOptions: IDropdownOption[] = [];

                        var i = 0;
                        if (sortedMessages !== undefined) {
                            for (const message of sortedMessages) {
                                if (i > 0)
                                    incidentMessagesDropdownOptions.push({
                                        key: message.createdDateTime,
                                        text: (new Date(message.createdDateTime)).toLocaleString()
                                    });
                                i++;
                            }
                        }

                        this.setState({
                            incidentId: result.id,
                            incidentTitle: result.title,
                            service: result.affectedServices,
                            status: result.statusDisplayName,
                            userImpact: result.impactDescription,
                            latestMessage: sortedMessages[0].description.content,
                            allMessages: sortedMessages,
                            messagesDropdownOptions: incidentMessagesDropdownOptions,
                            issueType: result.classificationDisplayName,
                            origin: result.originDisplayName,
                            lastUpdated: new Date(result.lastModifiedDateTime),
                            startTime: new Date(result.startDateTime),
                            endTime: result.endTime ? new Date(result.endDateTime) : undefined,
                            publicComm: result.public,
                            comments: result.publishingComments,
                            newComments: result.publishingComments,
                            summary: result.summary,
                            workItem: result.task.taskId && result.task.taskId.trim() !== "" ? "<a href='" + result.task.taskUrl + "' target='_blank'>" + result.task.taskId + "</a>" : "",
                            loading: false,
                            canWriteMetadata: canWriteMetadata
                        });

                        if (this.props.onView)
                            this.props.onView(this.state.incidentId, true);

                        if (this.props.onUpdateParentTitle)
                            this.props.onUpdateParentTitle(result.title);

                        var viewPoint: any = result.serviceHealthHubViewpoint;
                        viewPoint.viewed = true;

                        this.setState({
                            serviceHealthHubViewpoint: viewPoint,
                            loading: false
                        });

                    }).catch((err) => {
                        this.setState({
                            error: err.message
                        });
                    });
                }).catch((err) => {
                    this.setState({
                        error: err.message
                    });
                });
        }
    }

    _onLoadCustomActions(actions: ICustomAction[]): void {
        this.setState({
            customActions: actions
        });
    }

    private _onLoadTags(id: string | undefined, component: string): void {
        if (!id)
            return;

        getAvailableTagDefinitions(component,
            (res: ITagDefinition[]): void => {
                var availableTags: ITag[] = []
                var selectedTags: ITag[] = []

                for (const t of res) {
                    availableTags.push({
                        key: t.tagId,
                        name: t.name
                    });
                }

                getTags(id, component,
                    (res: MSTag[]): void => {

                        for (const t of res) {
                            var name: string | undefined = availableTags.find((at: ITag) => at.key === t.tagId)?.name
                            selectedTags.push({
                                key: t.tagId,
                                name: name ? name : ""
                            });
                        }

                        this.setState({
                            availableTags: availableTags
                        });
                    },
                    (message: string): void => {
                        var err: string | undefined = this.state.tagError;

                        if (!err)
                            err = message
                        else
                            err += " " + message

                        this.setState({
                            tagError: err
                        });
                    });


                this.setState({
                    availableTags: availableTags,
                    selectedTags: selectedTags
                });
            },
            (message: string): void => {
                var err: string | undefined = this.state.tagError;

                if (!err)
                    err = message
                else
                    err += " " + message

                this.setState({
                    tagError: err
                });
            });
    }

    private _onChangeOrgTags(items: ITag[] | undefined): void {
        var removedTags: ITag[] | undefined = undefined;
        var addedTags: ITag[] | undefined = undefined;

        if (!items) {
            removedTags = this.state.selectedTags;
        }
        else {
            addedTags = items.filter((t: ITag) => !this.state.selectedTags.includes(t));
            removedTags = this.state.selectedTags.filter((t: ITag) => !items.includes(t));

            for (const t of addedTags)
                addTag(this.state.incidentId, componentName, t.key.toString(),
                    (messageId: string, type: string, tagId: string): void => { },
                    (message: string): void => {
                        this.setState({
                            tagError: message
                        });
                    });

            for (const t of removedTags)
                removeTag(this.state.incidentId, componentName, t.key.toString(),
                    (messageId: string, type: string, tagId: string): void => { },
                    (message: string): void => {
                        this.setState({
                            tagError: message
                        });
                    });

            if (this.props.onUpdateOrgTags)
                this.props.onUpdateOrgTags(this.state.incidentId, items.map((i: ITag) => i.name));

            this.setState({
                selectedTags: items
            });
        }
    }

    private _setViewState(id: string, state: boolean): void {
        setViewState(id, state,
            (id: string, state: boolean) => {
                var viewpoint = this.state.serviceHealthHubViewpoint;
                viewpoint.viewed = state;
                this.setState({
                    serviceHealthHubViewpoint: viewpoint
                });

                if (this.props.onView)
                    this.props.onView(id, state);
            },
            (message: string) => {
                this.setState({
                    publishingError: message
                });
            });
    }

    private _setFavoriteState(id: string, state: boolean): void {
        setFavoriteState(id, state,
            (id: string, state: boolean) => {
                var viewpoint = this.state.serviceHealthHubViewpoint;
                viewpoint.favorite = state;
                this.setState({
                    serviceHealthHubViewpoint: viewpoint
                });

                if (this.props.onFavorite)
                    this.props.onFavorite(id, state);
            },
            (message: string) => {
                this.setState({
                    publishingError: message
                });
            });
    }

    private _setArchiveState(id: string, state: boolean): void {
        setArchiveState(id, state,
            (id: string, state: boolean) => {
                var viewpoint = this.state.serviceHealthHubViewpoint;
                viewpoint.archived = state;
                this.setState({
                    serviceHealthHubViewpoint: viewpoint
                });

                if (this.props.onArchive)
                    this.props.onArchive(id, state);
            },
            (message: string) => {
                this.setState({
                    publishingError: message
                });
            });
    }

    private _procecssPublishing = (): void => {
        var params: RequestInit;
        acquireAccessToken()
            .then((response) => {
                const body = {
                    id: this.state.incidentId,
                    type: componentName,
                    comments: this.state.newComments
                };

                params = {
                    headers: {
                        "Content-Type": "application/json charset=UTF-8",
                        "Authorization": "Bearer " + response.idToken
                    },
                    body: JSON.stringify(body),
                    method: "POST"
                };

                var responseObj: any = {
                    ok: false,
                    statusCode: null,
                    body: null
                };

                fetch("/api/PublicMessage/Publish", params)
                    .then((response) => {
                        responseObj.ok = response.ok;
                        responseObj.statusCode = response.status;
                        return response.ok ? null : response.text();
                    })
                    .then((body) => {
                        responseObj.body = body;

                        if (!responseObj.ok) {
                            // make the promise be rejected if we didn't get a 2xx response
                            const err = new Error("Couldn't publish message. Error details:<br/><br/><b>HTTP " + responseObj.statusCode + "</b><br/>" + body);
                            throw err;
                        }

                        this.setState({
                            comments: this.state.newComments,
                            publicComm: true
                        });

                        if (this.props.onPublishingChange)
                            this.props.onPublishingChange(this.state.incidentId, this.state.publicComm);

                        this._onDismisPublishMessageDialog();
                    })
                    .catch((err) => {
                        this._onDismisPublishMessageDialog();
                        this.setState({
                            publishingError: err.message
                        });
                    });
            })
            .catch((err) => {
                this._onDismisPublishMessageDialog();
                this.setState({
                    publishingError: "Authentication error. Details:<br/><br/>" + err.message
                });
            });;
    }

    private _procecssUnpublishing = (): void => {
        var params: RequestInit;
        acquireAccessToken()
            .then((response) => {
                const body = {
                    id: this.state.incidentId,
                    type: componentName
                };

                params = {
                    headers: {
                        "Content-Type": "application/json charset=UTF-8",
                        "Authorization": "Bearer " + response.idToken
                    },
                    body: JSON.stringify(body),
                    method: "POST"
                };

                var responseObj: any = {
                    ok: false,
                    statusCode: null,
                    body: null
                };

                fetch("/api/PublicMessage/Unpublish", params)
                    .then((response) => {
                        responseObj.ok = response.ok;
                        responseObj.statusCode = response.status;
                        return response.ok ? null : response.text();
                    })
                    .then((body) => {
                        responseObj.body = body;

                        if (!responseObj.ok) {
                            // make the promise be rejected if we didn't get a 2xx response
                            const err = new Error("Couldn't publish message. Error details:<br/><br/><b>HTTP " + responseObj.statusCode + "</b><br/>" + body);
                            throw err;
                        }

                        this.setState({
                            publicComm: false
                        });

                        if (this.props.onPublishingChange)
                            this.props.onPublishingChange(this.state.incidentId, this.state.publicComm);
                    })
                    .catch((err) => {
                        this.setState({
                            publishingError: err.message
                        });
                    });
            })
            .catch((err) => {
                this.setState({
                    publishingError: "Authentication error. Details:<br/><br/>" + err.message
                });
            });
    }

    private _cancelPublish = (): void => {
        this.setState({
            newComments: this.state.comments
        });

        this._onDismisPublishMessageDialog();
    }

    private _onOpenPublishMessageDialog = (): void => {
        this.setState({
            isPublishMessageDialogOpen: true
        });
    }

    private _onDismisPublishMessageDialog = (): void => {
        this.setState({
            isPublishMessageDialogOpen: false
        });
    }

    private _onPublicCommentsChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
        this.setState({
            newComments: newValue ? newValue : ""
        });
    }

    private _closePublishingErrorDialog = (): void => {
        this.setState({
            publishingError: undefined
        });
    }

    private _closeErrorDialog = (): void => {
        this.setState({
            error: undefined
        });
    }
}