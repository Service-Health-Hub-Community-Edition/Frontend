import * as React from 'react';
import {
    CommandBar, ICommandBarItemProps, IContextualMenuItem, Text, Icon, Pivot, PivotItem, FontIcon,
    DefaultPalette, mergeStyles, mergeStyleSets, ITag, TagPicker, IBasePickerSuggestionsProps
} from '@fluentui/react';
import { Separator } from '@fluentui/react';
import { Dropdown, IDropdownOption, IDropdownStyles, Button, ScrollablePane } from '@fluentui/react';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { MessageBar, MessageBarType, TextField, Dialog, DialogType, DialogFooter, PrimaryButton, TooltipHost } from '@fluentui/react';
import { MetaDataList, MetaDataItem } from '@m365-admin/metadata';
import { ICustomAction, CustomAction } from '../../../CustomAction';
import { AuditLog } from '../../../AuditLog';
import { Tag } from '../../../TagComponent';
import { AlertMessage } from '../../../AlertMessage';
import { acquireAccessToken } from "../../../../auth/AccessTokenHelper";
import { Communication } from '@microsoft/teams-js';
import { setViewState, setArchiveState, setFavoriteState } from "../../../../api/viewpoint";
import { getAvailableTagDefinitions, getTags, addTag, removeTag, MSTag } from '../../../../api/tags';
import { ITagDefinition } from '../../../admin/applicationSettings/TagDefinitions';

export interface IIncidentDetailsState {
    incidentId: string;
    incidentTitle: string;
    service: string;
    status: string;
    stage?: string;
    severity?: string;
    level?: number;
    incidentType?: string;
    region?: string;
    communicationId?: string;
    latestMessage: string;
    lastUpdated: Date;
    startTime: Date;
    endTime?: Date;
    workItem: string;
    summary?: any; 
    serviceHealthHubViewpoint: any;
    canWriteMetadata: boolean;
    selectedTags: ITag[];
    availableTags: ITag[];
    tagError?: string;
    customActions: ICustomAction[] | undefined;
    loading: boolean;
    error?: string;
}


const iconClass = mergeStyles({

});

const classNames = mergeStyleSets({
    incident: [{ color: 'red' }, iconClass],
    advisory: [{ color: DefaultPalette.blue }, iconClass]
});

const componentName: string = 'AzureServiceHealthAlert';

export class IncidentDetails extends React.Component<{
    id?: string,
    onPublishingChange?: any,
    onView?: any,
    onFavorite?: any,
    onArchive?: any,
    onUpdateOrgTags?: any
}, IIncidentDetailsState> {

    customAction: any = React.createRef();
    activityLog: any = React.createRef();

    constructor(props: {
        id?: string,
        onPublishingChange?: any,
        onView?: any,
        onFavorite?: any,
        onArchive?: any,
        onUpdateOrgTags?: any
    }) {
        super(props);

        this.state = {
            incidentId: "",
            incidentTitle: "",
            service: "",
            status: "",
            latestMessage: "",
            lastUpdated: new Date(),
            startTime: new Date(),
            endTime: undefined,
            workItem: "",
            serviceHealthHubViewpoint: undefined,
            canWriteMetadata: false,
            selectedTags: [],
            availableTags: [],
            customActions: undefined,
            loading: false,
            error: undefined
        };
    }

    public render() {
        const {
            incidentId, service, status, latestMessage, stage, severity, level, incidentType, region, communicationId, customActions,
            lastUpdated, startTime, workItem, summary, serviceHealthHubViewpoint, loading, error,
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
                                                {incidentId} · Start {startTime.toLocaleDateString([], { year: 'numeric', month: 'short', day: 'numeric' })} · Last updated {lastUpdated.toLocaleDateString([], { year: 'numeric', month: 'short', day: 'numeric' })}
                                            </Text>
                                        </div>
                                    </div>
                                    <div className="row">
                                        <div className="col">
                                            <CommandBar
                                                items={_items}
                                                farItems={[]}
                                                ariaLabel="Azure Service Health Alert actions"
                                            />
                                        </div>
                                    </div>
                                    <div className="row" style={{ paddingTop: "10px", marginBottom: "20px" }}>
                                        <div className="col">
                                            <MetaDataList>
                                                <MetaDataItem header='Service' body={ service !== undefined ? (<Tag name={service} />) : ""} />
                                                <MetaDataItem header='Task' body={workItem !== "" ? (<div><Icon iconName='TaskSolid' />&nbsp;<Text variant={'small'} >
                                                    <div style={{ whiteSpace: "pre-wrap", display: 'inline-block' }} dangerouslySetInnerHTML={{ __html: workItem }} />
                                                </Text></div>) : (<Text variant={'small'} >not present</Text>)} />
                                                <MetaDataItem header='Status' body={status} />
                                                <MetaDataItem header='Stage' body={stage} />
                                                <MetaDataItem header='Severity' body={
                                                    <div className="container" style={{ margin: "0px" }} >
                                                        <div className="row" style={{ paddingBottom: "0px" }} >
                                                            <div className="col-auto" style={{ margin: "0px", padding: "0px 6px 0px 0px" }} >
                                                                {level ? (<FontIcon
                                                                    iconName={level <= 1 ? 'WarningSolid' : 'InfoSolid'}
                                                                    className={level <= 1 ? classNames.incident : classNames.advisory} />) : ""}
                                                            </div>
                                                            <div className="col" style={{ margin: "0px", padding: "0px", alignItems: 'center' }}>
                                                                <Text variant='medium'>{severity}</Text>
                                                            </div>
                                                        </div>
                                                    </div>
                                                    } />
                                                <MetaDataItem header='Incident type' body={incidentType} />
                                                <MetaDataItem header='Region' body={region} />
                                                <MetaDataItem header='Communication Id' body={communicationId} />
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
                                    
                                    <div className="row">
                                        <div className="col">
                                            <Separator />
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
                                                <ul style={{ paddingTop: "6px", marginBottom: "0px" }} >
                                                    {summary.contents.map((summaryLine: any) => (
                                                        <li style={{ paddingTop: "6px" }}><Text variant='medium'>{summaryLine.text}</Text></li>
                                                    ))}
                                                </ul>
                                            </>
                                        ) : "" }
                                        <Text variant={'medium'} block>
                                            <br /><b>Latest message</b>
                                        </Text>
                                        <Text>
                                            <div style={{ whiteSpace: "pre-wrap" }} dangerouslySetInnerHTML={{ __html: latestMessage }} />
                                        </Text>
                                    </PivotItem>

                                    <PivotItem headerText="Activity">
                                        &nbsp;<br />
                                        {incidentId ? (<AuditLog itemId={incidentId} itemType={componentName} scheme='item' />) : (<Text variant='medium'>No activities available</Text>)}
                                    </PivotItem>
                                </Pivot>
                            </div>

                            <AlertMessage title='' header='Something went wrong.' message={error} isBlocking={true} isOpen={error !== undefined} onClose={this._closeErrorDialog} />
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

                    fetch('/api/azure/health/alerts/' + incidentId, { headers: response.idToken === "" ? {} : { 'Authorization': `Bearer ${response.idToken}` } })
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
                        var data: any = result.data ? JSON.parse(result.data)?.data : {};

                        this.setState({
                            incidentId: result.id,
                            incidentTitle: data.alertContext.properties.title,
                            service: data.alertContext.properties.service,
                            status: data.alertContext.status,
                            stage: data.alertContext.properties.stage,
                            level: data.alertContext.level,
                            severity: data.essentials.severity,
                            incidentType: data.alertContext.properties.incidentType,
                            region: data.alertContext.properties.region,
                            communicationId: data.alertContext.properties.communicationId,
                            latestMessage: data.alertContext.properties.communication,
                            lastUpdated: new Date(result.lastUpdatedTime),
                            startTime: new Date(data.alertContext.properties.impactStartTime),
                            endTime: data.alertContext.properties.impactMitigationTime ? new Date(data.alertContext.properties.impactMitigationTime) : undefined,
                            summary: result.additionalData.summary,
                            workItem: result.workItemID ? "<a href='" + result.workItemURL + "' target='_blank'>" + result.workItemID + "</a>" : "",
                            canWriteMetadata: canWriteMetadata
                        });

                        if (this.props.onView)
                            this.props.onView(this.state.incidentId, true);

                        var viewPoint: any = result.serviceHealthHubViewpoint;
                        viewPoint.viewed = true;

                        this.setState({
                            serviceHealthHubViewpoint: viewPoint,
                            loading: false
                        });
                    }).catch((err) => {
                        this.setState({
                            error: err.message,
                            loading: false
                        });
                    });
                }).catch((err) => {
                    this.setState({
                        error: err.message,
                        loading: false
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
                    error: message
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
                    error: message
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
                    error: message
                });
            });
    }

    private _closeErrorDialog = (): void => {
        this.setState({
            error: undefined
        });
    }
}