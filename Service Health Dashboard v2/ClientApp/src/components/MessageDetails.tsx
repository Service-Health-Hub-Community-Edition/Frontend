import * as React from 'react';
import { CommandBar, ICommandBarItemProps, IContextualMenuItem, Text, Icon, FontIcon } from '@fluentui/react';
import { Separator } from '@fluentui/react';
import { Spinner, SpinnerSize, ProgressIndicator } from '@fluentui/react';
import {
    Button, PrimaryButton, MessageBar, MessageBarType, Dialog, DialogType, DialogFooter, TextField, Pivot, PivotItem,
    TooltipHost, DetailsList, DetailsListLayoutMode, SelectionMode, ConstrainMode, ITag, TagPicker, IBasePickerSuggestionsProps
} from '@fluentui/react';
import { ServiceComponent } from './ServiceNameComponent';
import { Tag, TagViewState } from './TagComponent';
import { ICustomAction, CustomAction } from './CustomAction';
import { AuditLog } from './AuditLog';
import { acquireAccessToken } from "../auth/AccessTokenHelper";
import { AlertMessage } from './AlertMessage';
import { AuthenticationResult } from '@azure/msal-browser';
import { setViewState, setFavoriteState, setArchiveState } from '../api/viewpoint';
import { getAvailableTagDefinitions, getTags, addTag, removeTag, MSTag } from '../api/tags';
import { ITagDefinition } from './admin/applicationSettings/TagDefinitions';

interface IServiceMapping {
    service: string;
    serviceAliases: string[];
}

const serviceMapping: IServiceMapping[] = [
    {
        service: 'Exchange',
        serviceAliases: ['Exchange Online']
    },
    {
        service: 'Office 365',
        serviceAliases: ['Microsoft 365 Apps']
    },
    {
        service: 'OneDrive',
        serviceAliases: ['OneDrive for Business']
    },
    {
        service: 'SharePoint',
        serviceAliases: ['SharePoint Online']
    },
    {
        service: 'Skype For Business',
        serviceAliases: ['Skype for Business']
    },
    {
        service: 'Teams',
        serviceAliases: ['Microsoft Teams']
    },
    {
        service: 'Yammer',
        serviceAliases: ['Yammer']
    },
]

interface IServiceInfo {
    name: string;
    activeUsers: number;
    inactiveUsers: number;
    reportDate?: Date;
}

interface IMessageLink {
    name: string;
    url: string;
}

export interface IMessageDetailsState {
    id: string;
    title: string;
    service: IServiceInfo[];
    tag: string[];
    latestMessage: string;
    published: Date;
    lastUpdated: Date;
    actionRequiredBy?: Date;
    workItem: string;
    isMajorChange: boolean;
    links: IMessageLink[];
    publicComm: boolean;
    comments: string;
    newComments: string;
    rolloutStatus?: string;
    platforms?: string | string[];
    featureReleaseState: IFeatureReleaseState[];
    summary?: any;
    extendedProperties: any;
    isPublishMessageDialogOpen: boolean;
    category: string;
    publishingError?: string;
    error?: string;
    tagError?: string;
    initialized: boolean;
    serviceHealthHubViewpoint: any;
    customActions: ICustomAction[] | undefined;
    canWriteMetadata: boolean;
    selectedTags: ITag[];
    availableTags: ITag[];
}

interface IFeatureReleaseState {
    roadmapId: number;
    platform: string;
    status: string;
    latestRing: string | null;
    lastUpdateTime: string;
}

const componentName: string = 'ServiceUpdateMessage';

export class MessageDetails extends React.Component<{
    id?: string,
    onPublishingChange?: any,
    onExtendedPropertiesChange?: any,
    onView?: any,
    onFavorite?: any,
    onArchive?: any,
    onUpdateOrgTags?: any,
    onUpdateParentTitle?: any
}, IMessageDetailsState> {
    customAction: any = React.createRef();
    activityLog: any = React.createRef();

    constructor(props: {
        id?: string,
        onPublishingChange?: any,
        customAction?: any,
        onExtendedPropertiesChange?: any,
        onView?: any,
        onFavorite?: any,
        onArchive?: any,
        onUpdateOrgTags?: any,
        onUpdateParentTitle?: any
    }) {
        super(props);

        this.state = {
            id: "",
            title: "",
            service: [],
            tag: [],
            latestMessage: "",
            lastUpdated: new Date(),
            published: new Date(),
            actionRequiredBy: undefined,
            workItem: "",
            isMajorChange: false,
            category: "",
            links: [],
            publicComm: false,
            comments: "",
            newComments: "",
            featureReleaseState: [],
            extendedProperties: undefined,
            isPublishMessageDialogOpen: false,
            publishingError: undefined,
            error: undefined,
            initialized: false,
            serviceHealthHubViewpoint: undefined,
            customActions: undefined,
            canWriteMetadata: false,
            selectedTags: [],
            availableTags: []
        };
    }

    public render() {
        const {
            id, service, tag, latestMessage, published, lastUpdated,
            workItem, actionRequiredBy, category, isMajorChange, links, customActions,
            publicComm, newComments, serviceHealthHubViewpoint, summary, rolloutStatus, platforms,
            featureReleaseState, isPublishMessageDialogOpen, error, publishingError, initialized,
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
                        onClick: () => this.customAction.current._onRunCustomAction(customAction.name, customAction.actionId, id),
                    }) : 
                    customActionItems.push({
                        key: customAction.actionId,
                        text: customAction.name,
                        onClick: () => this.customAction.current._onRunCustomAction(customAction.name, customAction.actionId, id),
                    })
            }

        // add archive / restore command
        if (serviceHealthHubViewpoint)
            _items.push({
                key: 'archive',
                text: serviceHealthHubViewpoint.archived ? 'Restore' : 'Archive',
                iconProps: { iconName: serviceHealthHubViewpoint.archived ? 'Undo' : 'Archive' },
                onClick: () => this._setArchiveState(id, !serviceHealthHubViewpoint.archived),
            });

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
                    onClick: () => this._setViewState(id, false)
                } : {
                    key: 'markAsRead',
                    text: 'Mark as read',
                    iconProps: { iconName: 'Read' },
                    onClick: () => this._setViewState(id, true)
                });

            _items.push(serviceHealthHubViewpoint.favorite ?
                {
                    key: 'removeFavorite',
                    text: 'Remove from favorites',
                    iconProps: { iconName: 'FavoriteStarFill' },
                    onClick: () => this._setFavoriteState(id, false)
                } : {
                    key: 'addFavorite',
                    text: 'Add to favorites',
                    iconProps: { iconName: 'FavoriteStar' },
                    onClick: () => this._setFavoriteState(id, true)
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

                {initialized ? (
                    <div>
                        <div className="messageDetails" style={{ display: id ? 'block' : 'none' }}>                         
                            <div className="container" style={{ padding: '0px' }}>
                                <div className="row">
                                    <div className="col">
                                        <Text variant={'medium'} >
                                            {id} · Published {published.toLocaleDateString([], { year: 'numeric', month: 'short', day: 'numeric' })} · Last updated {lastUpdated.toLocaleDateString([], { year: 'numeric', month: 'short', day: 'numeric' })}
                                        </Text>
                                    </div>                                 
                                </div>
                                <div className="row">
                                    <div className="col">
                                        <CommandBar
                                            items={_items}
                                            farItems={[]}
                                            ariaLabel="Message Center communication actions"                                          
                                        />
                                    </div>
                                </div>
                                <div className="row" style={{ display: actionRequiredBy ? 'block' : 'none' }}>
                                    <div className="col">
                                        <MessageBar
                                            messageBarType={MessageBarType.severeWarning}
                                            isMultiline={false}
                                        >
                                            Action required by {actionRequiredBy?.toLocaleDateString()}
                                        </MessageBar>
                                    </div>
                                </div>
                                <div className="row" style={{ paddingTop: "10px", marginBottom: "20px" }}>
                                    <div className="col">
                                        <Text variant={'medium'} block>
                                            <b>Services & monthly active users</b>
                                        </Text>
                                        {service.map((s) => s.activeUsers >= 0 ? (<ServiceComponent name={s.name} additionalInfo={s.activeUsers.toString()} />) : (<ServiceComponent name={s.name} />))}
                                    </div>
                                    <div className="col">
                                        <Text variant={'medium'} block>
                                            <b>Task</b>
                                        </Text>

                                        {workItem !== "" ? (<div><FontIcon iconName='TaskSolid'/>&nbsp;<Text variant={'small'} >
                                            <div style={{ whiteSpace: "pre-wrap", display: 'inline-block' }} dangerouslySetInnerHTML={{ __html: workItem }} />
                                        </Text></div>) : (<Text variant={'small'} >not present</Text>)
                                        }
                                    </div>
                                </div>
                                
                                <div className="row" style={{ marginBottom: "20px" }}>
                                    <div className="col">
                                        {(category !== null && category.length > 0) || isMajorChange ? (
                                            <><Text variant={'medium'} block>
                                                <b>Category</b>
                                            </Text>
                                                <Text>
                                                    {category !== null && category.length > 0 ? (<Tag name={category} compact={false} />) : ""}
                                                    {isMajorChange ? (<Tag name={"Major change"} compact={false} />) : ""}
                                                </Text></>) : ""}
                                    </div>
                                    <div className="col">
                                        {tag.length > 0 ? (
                                            <Text variant={'medium'} block>
                                                <b>Tag</b>
                                            </Text>) : ""}

                                        {tag.map((t) => (
                                            <Tag name={t} compact={false} />
                                        ))}
                                        
                                    </div>
                                </div>  

                                {availableTags && availableTags.length > 0 ? (
                                    <div className="row" style={{ marginBottom: "20px" }}>
                                        <div className="col">
                                            <Text variant={'medium'} block>
                                                    <b>Organization tags</b>
                                            </Text>
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
                                            </MessageBar>) : (<></>) }
                                        </div>
                                    </div>) : (<></>)}

                                <div className="row">
                                    {(rolloutStatus !== undefined) ? (
                                    <div className="col">
                                        
                                            <><Text variant={'medium'} block>
                                                <b>Status for your org</b>
                                            </Text>
                                                <Text variant={'small'}>
                                                    {rolloutStatus}
                                            </Text><br/>&nbsp;
                                            </>
                                        </div>) : (<></>)}
                                    {(platforms !== undefined) ? (
                                        <div className="col">
                                        
                                            <><Text variant={'medium'} block>
                                                <b>Platforms</b>
                                            </Text>
                                                <Text variant={'small'}>
                                                    {Array.isArray(platforms) ? platforms.join(', ') : platforms }
                                                </Text><br />&nbsp;
                                            </>
                                        </div>) : (<></>)}
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
                                                <Icon iconName='Info' style={{ paddingLeft: '6px', cursor: 'default', fontSize: 'smaller' }} />
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
                                                </ul>) : (<div><br/>
                                                <Text variant='medium'>{summary.contents[0].text}</Text><br/><br/>&nbsp;</div>
                                            )}
                                            <div style={{ marginBottom: "6px" }}>
                                                <Text variant='medium'><b>Details</b></Text>
                                            </div>
                                        </>
                                    ) : ""}

                                    <Text>
                                        <div dangerouslySetInnerHTML={{ __html: latestMessage }} />
                                    </Text>

                                    <div>
                                        {links.map((l: any) => (<Button text={l.name} href={l.url} target="_blank" styles={{ root: { margin: "8px" } }} />))}
                                    </div>
                                    <Separator />
                                </PivotItem>

                                <PivotItem headerText="Activity">
                                    &nbsp;<br />
                                    {id ? (<AuditLog itemId={id} itemType={componentName} scheme='item' ref={this.activityLog} />) : (<Text variant='medium'>No activities available</Text>)}
                                </PivotItem>

                                {featureReleaseState && featureReleaseState.length > 0 ? (
                                    <PivotItem headerText="Release status">
                                    <DetailsList
                                        items={featureReleaseState}
                                        compact={true}
                                        columns={[
                                            {
                                                key: 'roadmapId',
                                                name: 'Roadmap Id',
                                                minWidth: 80,
                                                maxWidth: 80,
                                                isResizable: false,
                                                isCollapsible: true,
                                                isPadded: false,
                                                onRender: (item: IFeatureReleaseState) => {
                                                    return (<>
                                                        {item.roadmapId}
                                                    </>)
                                                },
                                            },
                                            {
                                                key: 'platform',
                                                name: 'Platform',
                                                minWidth: 80,
                                                maxWidth: 80,
                                                isResizable: false,
                                                isCollapsible: true,
                                                isPadded: false,
                                                onRender: (item: IFeatureReleaseState) => {
                                                    return (<>
                                                        {item.platform}
                                                    </>)
                                                },
                                            },
                                            {
                                                key: 'status',
                                                name: 'Status',
                                                minWidth: 64,
                                                maxWidth: 64,
                                                isResizable: false,
                                                isCollapsible: true,
                                                isPadded: false,
                                                onRender: (item: IFeatureReleaseState) => {
                                                    return (<>
                                                        {item.status}
                                                    </>)
                                                },
                                            },
                                            {
                                                key: 'latestRing',
                                                name: 'Ring',
                                                minWidth: 100,
                                                maxWidth: 100,
                                                isResizable: false,
                                                isCollapsible: true,
                                                isPadded: false,
                                                onRender: (item: IFeatureReleaseState) => {
                                                    return (<>
                                                        {item.latestRing !== null && item.latestRing !== undefined ? item.latestRing : ""}
                                                    </>)
                                                },
                                            },
                                            {
                                                key: 'updated',
                                                name: 'Updated',
                                                minWidth: 64,
                                                maxWidth: 64,
                                                isResizable: false,
                                                isCollapsible: true,
                                                isPadded: false,
                                                onRender: (item: IFeatureReleaseState) => {
                                                    return (<>
                                                        {item.lastUpdateTime !== null && item.lastUpdateTime !== undefined ? new Date(item.lastUpdateTime).toLocaleDateString() : ""}
                                                    </>)
                                                },
                                            },
                                        ]}
                                        styles={{
                                            contentWrapper: {
                                                overflowX: 'hidden'
                                            },
                                            root: {
                                                overflowX: 'hidden'
                                            }
                                        }}
                                        selectionMode={SelectionMode.none}
                                        layoutMode={DetailsListLayoutMode.justified}
                                        constrainMode={ConstrainMode.horizontalConstrained}
                                        isHeaderVisible={true}
                                        />
                                    </PivotItem>
                                ) : (<></>)}
                            </Pivot>                  
                        </div>

                        <Dialog
                            hidden={!isPublishMessageDialogOpen}
                            onDismiss={this._onDismisPublishMessageDialog}
                            dialogContentProps={{
                                type: DialogType.normal,
                                title: 'Publish announcement',
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
                                multiline rows={8}
                                onChange={this._onPublicCommentsChange} />

                            <DialogFooter>
                                <PrimaryButton onClick={this._procecssPublishing} text="Save" />&nbsp;
                                <Button onClick={this._cancelPublish} text="Cancel" />
                            </DialogFooter>
                        </Dialog>

                        <AlertMessage title='' header='Something went wrong.' message={error} isBlocking={true} isOpen={error !== undefined} onClose={this._closeErrorDialog} />
                        <AlertMessage title='' header='Something went wrong.' message={publishingError} isBlocking={true} isOpen={publishingError !== undefined} onClose={this._closePublishingErrorDialog} />

                    </div>) : (
                    <div className="loadingProgress" style={{ display: id && error !== undefined ? 'none' : '' }}>
                        <br /><br />
                        <Spinner size={SpinnerSize.large} />
                    </div>)}
            </div>
        );
    }

    componentDidMount() {
        this._onLoadTags(this.props.id, componentName);
        this._getMessageDetails(this.props.id);
    }

    componentDidUpdate(prevProps: { id?: string }) {
        if (prevProps.id !== this.props.id) {
            this.setState({
                initialized: false
            });

            this._onLoadTags(this.props.id, componentName);
            this._getMessageDetails(this.props.id);
        }
    }

    _onLoadCustomActions(actions: ICustomAction[]): void {
        this.setState({
            customActions: actions
        });
    }

    private _getMessageDetails(id?: string) {
        const requiredRoles: string[] = ['ServiceHealthReader', 'Communication.Write.All', 'Admin'];
        const commsMgrRoles: string[] = ['Communication.Write.All', 'Admin'];
        var authResponse: AuthenticationResult;
        var mauCollection: IServiceInfo[] = [];
        var messageId = id;

        if (messageId !== undefined && messageId.trim().length > 0) {
            acquireAccessToken()
                .then((response) => {
                    var tokenClaims: any = response.account?.idTokenClaims;
                    const userRoles: any = tokenClaims?.roles;
                    var userHasRequiredRole: boolean = userRoles.some((r: string) => requiredRoles.includes(r));
                    var canWriteMetadata: boolean = userRoles.some((r: string) => commsMgrRoles.includes(r));

                    authResponse = response;

                    if (userHasRequiredRole)
                        fetch('/api/MonthlyActiveUsers', { headers: authResponse.idToken === "" ? {} : { 'Authorization': `Bearer ${authResponse.idToken}` } })
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
                                for (const mauStat of result)
                                    mauCollection.push({
                                        name: mauStat.service,
                                        activeUsers: mauStat.active,
                                        inactiveUsers: mauStat.inactive,
                                        reportDate: mauStat.reportDate
                                    });
                            })
                            .then(() =>
                                fetch('/api/Messages/' + messageId, { headers: authResponse.idToken === "" ? {} : { 'Authorization': `Bearer ${authResponse.idToken}` } })
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
                                        if (result !== null) {
                                            var serviceCollection: IServiceInfo[] = [];
                                            var serviceNames: string[] = result.services.length > 0 ? result.services : ["General announcement"];

                                            for (const service of serviceNames) {
                                                const mappedService: IServiceMapping | undefined = serviceMapping.find((s) => s.serviceAliases.includes(service));
                                                var serviceInfo: IServiceInfo | undefined = undefined;
                                                if (mappedService != undefined) {
                                                    serviceInfo = mauCollection.find((s) => s.name === mappedService.service);
                                                }

                                                serviceCollection.push({
                                                    name: service,
                                                    activeUsers: serviceInfo !== undefined ? serviceInfo.activeUsers : -1,
                                                    inactiveUsers: serviceInfo !== undefined ? serviceInfo.inactiveUsers : -1,
                                                    reportDate: serviceInfo !== undefined ? serviceInfo.reportDate : undefined
                                                });
                                            }

                                            var featureStatusJson = result.details?.find((d: any) => d.name === 'FeatureStatusJson');
                                            var rolloutStatus: string | undefined = undefined;
                                            var releaseStateDetails: IFeatureReleaseState[] = [];

                                            if (featureStatusJson !== undefined) {
                                                try {
                                                    var featureStatus = JSON.parse(featureStatusJson.value);
                                                    var fsKeys = Object.keys(featureStatus)
                                                    var globalStateRoadmap = featureStatus[fsKeys[0]].find((gsr: any) => gsr.Platform === 'All');

                                                    for (const roadmapId of fsKeys) {
                                                        var roadmapReleaseStates = featureStatus[roadmapId];
                                                        for (const roadmapReleaseState of roadmapReleaseStates) {
                                                            const rsState: string = roadmapReleaseState.Status;
                                                            if (roadmapReleaseState.Status !== 'FeatureRolloutStatusNotSupported')
                                                                releaseStateDetails.push({
                                                                    roadmapId: roadmapReleaseState.RoadmapId,
                                                                    platform: roadmapReleaseState.Platform,
                                                                    status: rsState.toLowerCase() === 'inrollout' ? 'Rolling out' : rsState,
                                                                    latestRing: roadmapReleaseState.LatestRing,
                                                                    lastUpdateTime: roadmapReleaseState.LastUpdateTime
                                                                });
                                                        }
                                                    }

                                                    rolloutStatus = globalStateRoadmap && globalStateRoadmap.Status !== 'FeatureRolloutStatusNotSupported' ? globalStateRoadmap.Status : undefined
                                                    if (rolloutStatus?.toLowerCase() === 'inrollout')
                                                        rolloutStatus = 'Rolling out'
                                                } catch (e) {
                                                    rolloutStatus = undefined
                                                }
                                            }

                                            var platformsList = result.details?.find((d: any) => d.name === 'Platforms');
                                            var platforms = platformsList ? platformsList.value?.split(', ') : undefined;

                                            var links: IMessageLink[] = [];

                                            const externalLink: any | undefined = result.details.find((d: any) => d.name === "ExternalLink");
                                            const helpLink: any | undefined = result.details.find((d: any) => d.name === "HelpLink");
                                            const blogLink: any | undefined = result.details.find((d: any) => d.name === "BlogLink");
                                            
                                            if (externalLink !== undefined)
                                                links.push({
                                                    name: "Additional information",
                                                    url: externalLink.value
                                                });

                                            if (helpLink !== undefined)
                                                links.push({
                                                    name: "Help and support",
                                                    url: helpLink.value
                                                });

                                            if (blogLink !== undefined)
                                                links.push({
                                                    name: "Blog",
                                                    url: blogLink.value
                                                });

                                            this.setState({
                                                id: result.id,
                                                title: result.title,
                                                service: serviceCollection,
                                                tag: result.tags,
                                                latestMessage: result.body.content.replaceAll(/\[(.*?)\]/g, '<b>$1</b>').replace("<p><br></p>", ""),
                                                lastUpdated: new Date(result.lastModifiedDateTime),
                                                published: new Date(result.startDateTime),
                                                actionRequiredBy: result.actionRequiredByDateTime ? new Date(result.actionRequiredByDateTime) : undefined,
                                                isMajorChange: result.isMajorChange,
                                                category: result.categoryDisplayName,
                                                links: links,
                                                rolloutStatus: rolloutStatus,
                                                featureReleaseState: releaseStateDetails,
                                                publicComm: result.public,
                                                comments: result.publishingComments,
                                                newComments: result.publishingComments,
                                                extendedProperties: result.extendedProperties,
                                                platforms: platforms,
                                                summary: result.summary,
                                                workItem: result.task.taskId && result.task.taskId.trim() !== "" ? "<a href='" + result.task.taskUrl + "' target='_blank'>" + result.task.taskId + "</a>" : "",
                                                canWriteMetadata: canWriteMetadata
                                            });

                                            if (this.props.onView)
                                                this.props.onView(this.state.id, true);

                                            if (this.props.onUpdateParentTitle)
                                                this.props.onUpdateParentTitle(result.title);

                                            var viewPoint: any = result.serviceHealthHubViewpoint;
                                            viewPoint.viewed = true;

                                            this.setState({
                                                serviceHealthHubViewpoint: viewPoint,
                                                initialized: true
                                            });
                                        }
                                    }));
                }).catch((err) => {
                    this.setState({
                        error: err.message
                    });
                });
        }
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
                addTag(this.state.id, componentName, t.key.toString(),
                    (messageId: string, type: string, tagId: string): void => { },
                    (message: string): void => {
                        this.setState({
                            tagError: message
                        });
                    });

            for (const t of removedTags)
                removeTag(this.state.id, componentName, t.key.toString(),
                    (messageId: string, type: string, tagId: string): void => { },
                    (message: string): void => {
                        this.setState({
                            tagError: message
                        });
                    });

            if (this.props.onUpdateOrgTags)
                this.props.onUpdateOrgTags(this.state.id, items.map((i: ITag) => i.name));

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
                    id: this.state.id,
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
                            this.props.onPublishingChange(this.state.id, this.state.publicComm);

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
                    id: this.state.id,
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
                            this.props.onPublishingChange(this.state.id, this.state.publicComm);
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

    private _onPublicCommentsChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
        this.setState({
            newComments: newValue ? newValue : ""
        });
    }

    private _setArchiveFlag = (archive: boolean): void => {
        var params: RequestInit;
        acquireAccessToken()
            .then((response) => {
                const body = {
                    id: this.state.id,
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

                const uri: string = archive ? '/api/Items/Archive' : '/api/Items/Restore';
                var responseObj: any = {
                    ok: false,
                    statusCode: null,
                    body: null
                };

                fetch(uri, params)
                    .then((response) => {
                        responseObj.ok = response.ok;
                        responseObj.statusCode = response.status;
                        return response.ok ? response.json() : response.text();
                    })
                    .then((res) => {
                        responseObj.body = res;

                        if (!responseObj.ok) {
                            // make the promise be rejected if we didn't get a 2xx response
                            const err = new Error("Couldn't " + (archive ? 'archive' : 'restore') + " message. Error details:<br/><br/><b>HTTP " + responseObj.statusCode + "</b><br/>" + responseObj.body);
                            throw err;
                        }

                        return responseObj.body;
                    })
                    .then((res) => {
                        if (res.result === 1) {
                            var extendedProperties = this.state.extendedProperties;
                            if (extendedProperties === undefined || extendedProperties === null)
                                extendedProperties = {
                                    Archived: archive
                                }
                            else
                                extendedProperties['Archived'] = archive;

                            this.setState({
                                extendedProperties: extendedProperties
                            });

                            if (this.props.onExtendedPropertiesChange)
                                this.props.onExtendedPropertiesChange(this.state.id, this.state.extendedProperties);

                            if (this.activityLog.current !== undefined && this.activityLog.current !== null)
                                this.activityLog.current._onLoadData();
                        } else {
                            var msg: string = "Couldn't " + (archive ? 'archive' : 'restore') + " message.";
                            if (res.result === -1)
                                msg += " Item is not found in the database. Please wait for synchronization to complete and try again. Synchronization can be confirmed within the activity log for the item.";
                            else if (res.result === 0)
                                msg += " Tried to " + (archive ? 'archive' : 'restore') + " already " + (archive ? 'archive' : 'restore') +"d message.";
                            const err = new Error(msg);
                            throw err;
                        }
                    })
                    .catch((err) => {
                        this.setState({
                            error: err.message
                        });
                    });
            });
    }
}