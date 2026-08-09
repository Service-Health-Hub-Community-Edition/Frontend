import * as React from 'react';
import {
    CommandBar, ICommandBarItemProps, IContextualMenuItem, Text, Icon, Pivot, PivotItem, FontIcon,
    DefaultPalette, mergeStyles, mergeStyleSets, ITag, TagPicker, IBasePickerSuggestionsProps
} from '@fluentui/react';
import { Separator } from '@fluentui/react';
import { Dropdown, IDropdownOption, IDropdownStyles, Button, DefaultButton, ScrollablePane } from '@fluentui/react';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { MessageBar, MessageBarType, TextField, Dialog, DialogType, DialogFooter, PrimaryButton, TooltipHost } from '@fluentui/react';
import { MetaDataList, MetaDataItem } from '@m365-admin/metadata';
import { ICustomAction, CustomAction } from '../../../CustomAction';
import { AuditLog } from '../../../AuditLog';
import { Tag } from '../../../TagComponent';
import { AlertMessage } from '../../../AlertMessage';
import { acquireAccessToken } from "../../../../auth/AccessTokenHelper";
import { setViewState, setArchiveState, setFavoriteState } from "../../../../api/viewpoint";
import { getAvailableTagDefinitions, getTags, addTag, removeTag, MSTag } from '../../../../api/tags';
import { ITagDefinition } from '../../../admin/applicationSettings/TagDefinitions';

export interface ID365PPReleaseDetailsState {
    id: string;
    title: string;
    businessValue: string;
    description: string;
    featureType: string;
    product: string;
    productArea: string;
    parentProduct: string;
    enabledFor: string;
    releaseWaveId?: string;
    releaseWave?: string;
    rwStartShipDate?: Date;
    rwEndShipDate?: Date;
    rwStatus?: string;
    earlyAccess: string;
    publicPreview: string;
    ga: string;
    documentation?: string;
    blogArticle?: string;
    overviewVideo?: string;
    published?: Date;
    lastUpdate?: Date;
    favorite: boolean;
    serviceHealthHubState: string;
    serviceHealthHubViewpoint: any;
    shhImageMetadata?: any;
    workItem: string;
    extendedProperties?: any;
    customActions: ICustomAction[] | undefined;
    selectedTags: ITag[];
    availableTags: ITag[];
    canWriteMetadata: boolean;
    tagError?: string;
    loading: boolean;
    error?: string;
}


const iconClass = mergeStyles({

});

const classNames = mergeStyleSets({
    incident: [{ color: 'red' }, iconClass],
    advisory: [{ color: DefaultPalette.blue }, iconClass]
});

const componentName: string = 'D365PowerPlatformRelease';

export class D365PPReleaseDetails extends React.Component<{
    id?: string,
    onPublishingChange?: any,
    onView?: any,
    onFavorite?: any,
    onArchive?: any,
    onUpdateOrgTags?: any,
    onUpdateParentTitle?: any
}, ID365PPReleaseDetailsState> {

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
            id: "",
            title: "",
            businessValue: "",
            description: "",
            featureType: "",
            product: "",
            productArea: "",
            parentProduct: "",
            enabledFor: "",
            earlyAccess: "-",
            publicPreview: "-",
            ga: "-",
            favorite: false,
            published: new Date(),
            serviceHealthHubState: "",
            workItem: "",
            serviceHealthHubViewpoint: undefined,
            customActions: undefined,
            selectedTags: [],
            availableTags: [],
            canWriteMetadata: false,
            loading: false,
            error: undefined
        };
    }

    public render() {
        const {
            id, title, businessValue, description, featureType, product, productArea, parentProduct, enabledFor, customActions, published, lastUpdate,
            releaseWave, rwStartShipDate, rwEndShipDate, rwStatus, shhImageMetadata, documentation, blogArticle, overviewVideo, earlyAccess, publicPreview, ga,
            workItem, serviceHealthHubViewpoint, loading, error, canWriteMetadata, selectedTags, availableTags, tagError
        } = this.state;

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

                {loading ? (
                    <div className="loadingProgress">
                        <br /><br />
                        <Spinner size={SpinnerSize.large} />
                    </div>
                ) :
                (
                        <div>
                            <div className="incidentDetails" style={{ display: id ? 'block' : 'none' }}>
                                <div className="container" style={{ padding: '0px' }}>
                                    <div className="row">
                                        <div className="col">
                                            <Text variant={'medium'} >
                                                Published {published ? published.toLocaleDateString([], { year: 'numeric', month: 'short', day: 'numeric' }) : '-'} · Last updated {lastUpdate ? lastUpdate.toLocaleDateString([], { year: 'numeric', month: 'short', day: 'numeric' }) : '-'}
                                            </Text>
                                        </div>
                                    </div>
                                    <div className="row">
                                        <div className="col">
                                            <CommandBar
                                                items={_items}
                                                farItems={[]}
                                                ariaLabel="Dynamics 365 and Power Platform Release actions"
                                            />
                                        </div>
                                    </div>
                                    <div className="row" style={{ paddingTop: "10px", marginBottom: "20px" }}>
                                        <div className="col">
                                            <MetaDataList>
                                                <MetaDataItem header='Product' body={product ? product : ""} />
                                                <MetaDataItem header='Product area' body={productArea ? productArea : ""} />
                                                <MetaDataItem header='Parent product' body={parentProduct ? parentProduct : ""} />
                                                <MetaDataItem header='Task' body={workItem !== "" ? (<div><Icon iconName='TaskSolid' />&nbsp;<Text variant={'small'} >
                                                    <div style={{ whiteSpace: "pre-wrap", display: 'inline-block' }} dangerouslySetInnerHTML={{ __html: workItem }} />
                                                </Text></div>) : (<Text variant={'small'} >not present</Text>)} />
                                                <MetaDataItem header='Feature type' body={featureType ? featureType : ""} />
                                                <MetaDataItem header='Enabled for' body={enabledFor ? enabledFor : ""} />

                                                <MetaDataItem header='Early access' body={earlyAccess} />
                                                <MetaDataItem header='Public preview' body={publicPreview} />
                                                <MetaDataItem header='General availability' body={ga} />

                                                <MetaDataItem header='Release wave' body={releaseWave ? releaseWave : ""} />

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
                                        <Text variant={'medium'} block>
                                            <br /><b>Business value</b>
                                        </Text>
                                        <Text>
                                            <div style={{ whiteSpace: "pre-wrap" }} dangerouslySetInnerHTML={{ __html: businessValue }} />
                                        </Text>
                                        &nbsp;<br />
                                        <Text variant={'medium'} block>
                                            <br /><b>Feature description</b>
                                        </Text>
                                        <Text>
                                            <div style={{ whiteSpace: "pre-wrap" }} dangerouslySetInnerHTML={{ __html: description }} />
                                        </Text>
                                        <div>
                                            {documentation ? (<PrimaryButton text="Documentation" href={documentation} target="_blank" styles={{ root: { margin: "8px" } }} />) : (<></>)}
                                            {blogArticle ? (<PrimaryButton text="Blog article" href={blogArticle} target="_blank" styles={{ root: { margin: "8px" } }} />) : (<></>)}
                                            {overviewVideo ? (<PrimaryButton text="Overview video" href={overviewVideo} target="_blank" styles={{ root: { margin: "8px" } }} />) : (<></>)}
                                        </div>
                                    </PivotItem>

                                    <PivotItem headerText="Activity">
                                        &nbsp;<br />
                                        {id ? (<AuditLog itemId={id} itemType={componentName} scheme='item' />) : (<Text variant='medium'>No activities available</Text>)}
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
        var updateId = id;
        if (updateId !== undefined) {
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

                    fetch('/api/d365pp/releases/' + updateId, { headers: response.idToken === "" ? {} : { 'Authorization': `Bearer ${response.idToken}` } })
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
                        const earlyAccess = this._getReleaseStatusString(result.earlyAccessStatus, result.earlyAccessDate);
                        const publicPreview = this._getReleaseStatusString(result.publicPreviewStatus, result.publicPreviewDate);
                        const ga = this._getReleaseStatusString(result.gaStatus, result.gaDate);

                        this.setState({
                            id: result.id,
                            title: result.title,
                            businessValue: result.businessValue,
                            description: result.description,
                            featureType: result.featureType,
                            product: result.product,
                            productArea: result.productArea,
                            parentProduct: result.parentProduct,
                            enabledFor: result.enabledFor,
                            releaseWaveId: result.releaseWaveId,
                            releaseWave: result.releaseWave,
                            rwStartShipDate: result.rwStartShipDate,
                            rwEndShipDate: result.rwEndShipDate,
                            rwStatus: result.rW_Status,
                            earlyAccess: earlyAccess,
                            publicPreview: publicPreview,
                            ga: ga,
                            documentation: result.documentation,
                            blogArticle: result.blogArticle,
                            overviewVideo: result.overviewVideo,
                            lastUpdate: result.lastUpdate ? new Date(result.lastUpdate) : undefined,
                            published: result.published ? new Date(result.published) : undefined,
                            shhImageMetadata: result.shhImageMetadata,
                            favorite: result.serviceHealthHubViewpoint?.favorited ? result.serviceHealthHubViewpoint.favorited : false,
                            serviceHealthHubState: result.state ? result.state : "",
                            serviceHealthHubViewpoint: result.serviceHealthHubViewpoint,
                            workItem: result.workItemID ? "<a href='" + result.workItemURL + "' target='_blank'>" + result.workItemID + "</a>" : "",
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

    _getReleaseStatusString(status: string | undefined, date: any | undefined) : string{
        var res = ""

        res = date ? new Date(date).toLocaleDateString([], { year: 'numeric', month: 'short', day: 'numeric' }) : ""

        if (status) {
            if (res !== "")
                res += " (" + status + ")"
            else
                res = status
        }

        if (res === "")
            res = "-";

        return res;
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