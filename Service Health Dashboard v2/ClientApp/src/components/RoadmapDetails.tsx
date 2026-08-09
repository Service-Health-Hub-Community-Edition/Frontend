import * as React from 'react';
import { CommandBar, ICommandBarItemProps, IContextualMenuItem, Text, Icon, FontIcon } from '@fluentui/react';
import { Separator } from '@fluentui/react';
import { Spinner, SpinnerSize, ProgressIndicator, Pivot, PivotItem } from '@fluentui/react';
import { Button, PrimaryButton, MessageBar, MessageBarType, Dialog, DialogType, DialogFooter, TextField } from '@fluentui/react';
import { ServiceComponent } from './ServiceNameComponent';
import { Tag, TagViewState } from './TagComponent';
import { ICustomAction, CustomAction } from './CustomAction';
import { AuditLog } from './AuditLog';
import { AlertMessage } from './AlertMessage';
import { acquireAccessToken } from "../auth/AccessTokenHelper";
import { AuthenticationResult } from '@azure/msal-browser';

export interface IRoadmapTag {
    tagName: string;
}

export interface IRoadmapItem {
    id: string;
    title: string;
    status: string;
    links: IMessageLink[];
    availabilityFrom: Date;
    availabilityTo: Date;
    availabilityDate: string;
    publicPreviewFrom: Date;
    publicPreviewTo: Date;
    publicPreviewDate: string;
    description: string;
    categories: string[];
    tags: string[];
    cloudInstances: string[];
    products: string[];
    releasePhase: string[];
    platforms: string[];
    published: Date;
    lastUpdated: Date;
    workItem: string;
    extendedProperties: any;
}

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

export interface IMessageLink {
    name: string;
    url: string;
}

export interface IRoadmapDetailsState {
    id: string;
    title: string;
    status: string;
    links: IMessageLink[];
    availabilityFrom: Date;
    availabilityTo: Date;
    availabilityDate: string;
    publicPreviewFrom: Date;
    publicPreviewTo: Date;
    publicPreviewDate: string;
    description: string;
    categories: string[];
    tags: string[];
    cloudInstances: string[];
    products: string[];
    releasePhase: string[];
    platforms: string[];
    published: Date;
    lastUpdated: Date;
    workItem: string;
    extendedProperties: any;
    error?: string;
    initialized: boolean;
}

const componentName: string = 'RoadmapCommunication';

export class RoadmapDetails extends React.Component<{ id?: string, onPublishingChange?: any, onExtendedPropertiesChange?: any, onUpdateParentTitle?: any }, IRoadmapDetailsState> {
    customAction: any = React.createRef();
    activityLog: any = React.createRef();

    constructor(props: { id?: string, onPublishingChange?: any, customAction?: any, onExtendedPropertiesChange?: any, onUpdateParentTitle?: any }) {
        super(props);

        this.state = {
            id: "",
            title: "",
            status: '',
            links: [],
            description: "",
            availabilityFrom: new Date(),
            availabilityTo: new Date(),
            availabilityDate: "",
            publicPreviewFrom: new Date(),
            publicPreviewTo: new Date(),
            publicPreviewDate: "",
            categories: [],
            tags: [],
            cloudInstances: [],
            products: [],
            releasePhase: [],
            platforms: [],
            lastUpdated: new Date(),
            published: new Date(),
            workItem: "",
            extendedProperties: undefined,
            error: undefined,
            initialized: false
        };
    }

    public render() {
        const {
            id, status, links,
            categories, tags, cloudInstances, products, releasePhase, platforms,
            description, published, lastUpdated, extendedProperties,
            workItem, availabilityDate, publicPreviewDate,
            error, initialized } = this.state;

        var _items: ICommandBarItemProps[] = [];
        var customActionItems: IContextualMenuItem[] = [];
        var customActions: ICustomAction[] = [];

        if (this.customAction.current !== undefined && this.customAction.current !== null)
            customActions = this.customAction.current._getCustomActions();

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
        if (extendedProperties?.Archived)
            _items.push({
                key: 'restore',
                text: 'Restore',
                iconProps: { iconName: 'Undo' },
                onClick: () => this._setArchiveFlag(false),
            });
        else
            _items.push({
                key: 'archive',
                text: 'Archive',
                iconProps: { iconName: 'Archive' },
                onClick: () => this._setArchiveFlag(true),
            });

        /*
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

        */

        // add custom actions
        if (customActionItems.length > 0)
            _items.push({
                key: "itemIntegrate",
                text: "Integrate",
                iconProps: { iconName: "AppIconDefault" },
                subMenuProps: { items: customActionItems }
            });   

        return (
            <div>
                <CustomAction componentName={componentName} ref={this.customAction} />

                {initialized ? (
                    <div>
                        <div className="roadmapDetails" style={{ display: id ? 'block' : 'none' }}>                         
                            <div className="container" style={{ padding: '0px' }}>
                                <div className="row">
                                    <div className="col">
                                        <Text variant={'medium'} >
                                            {id} · Published {published.toLocaleDateString([], { year: 'numeric', month: 'short', day: 'numeric' })} · Last updated {lastUpdated.toLocaleDateString([], { year: 'numeric', month: 'short', day: 'numeric' })}
                                        </Text>
                                    </div>                                 
                                </div>
                                <div className="row">
                                    {_items.length > 0 ? (<div className="col">
                                        <CommandBar
                                            items={_items}
                                            farItems={[]}
                                            ariaLabel="Roadmap communication actions"
                                        />
                                    </div>) : (<div />)}
                                </div>
                                <div className="row" style={{ paddingTop: "10px", marginBottom: "20px" }}>
                                    <div className="col">
                                        <Text variant={'medium'} block>
                                            <b>Public preview</b>
                                        </Text>

                                        <Text variant={'small'} >
                                            {publicPreviewDate}
                                        </Text>
                                    </div>
                                    <div className="col">
                                        <Text variant={'medium'} block>
                                            <b>General availability</b>
                                        </Text>

                                        <Text variant={'small'} >
                                            {availabilityDate}
                                        </Text>
                                    </div>                                   
                                </div>

                                {cloudInstances.length > 0 || releasePhase.length > 0 || platforms.length > 0 ? (<div>
                                    <div className="row" style={{ marginBottom: "20px" }}>
                                        <div className="col">
                                            <Text variant={'medium'} block>
                                                <b>Products</b>
                                            </Text>
                                            <Text>
                                                {products ? products.map((t) => (
                                                    <ServiceComponent name={t} />
                                                )) : ""}
                                            </Text>
                                        </div>
                                    </div>
                                    <div className="row" style={{ marginBottom: "20px" }}>
                                        <div className="col">
                                            <Text variant={'medium'} block>
                                                <b>Cloud instances</b>
                                            </Text>
                                            <Text>
                                                {cloudInstances ? cloudInstances.map((t) => (
                                                    <Tag name={t} compact={false} viewState={TagViewState.TextOnly} />
                                                )) : ""}
                                            </Text>
                                        </div>
                                        <div className="col">
                                            <Text variant={'medium'} block>
                                                <b>Platforms</b>
                                            </Text>
                                            <Text>
                                                {platforms ? platforms.map((t) => (
                                                    <Tag name={t} compact={false} viewState={TagViewState.TextOnly} />
                                                )) : ""}
                                            </Text>
                                        </div>
                                    </div>

                                    <div className="row" style={{ marginBottom: "20px" }}>
                                        <div className="col">
                                            <Text variant={'medium'} block>
                                                <b>Release phase</b>
                                            </Text>
                                            <Text>
                                                {releasePhase ? releasePhase.map((t) => (
                                                    <Tag name={t} compact={false} viewState={TagViewState.TextOnly} />
                                                )) : ""}
                                            </Text>
                                        </div>
                                        <div className="col">
                                            <Text variant={'medium'} block>
                                                <b>Task</b>
                                            </Text>

                                            {workItem !== "" ? (<div><FontIcon iconName='TaskSolid' />&nbsp;<Text variant={'small'} >
                                                <div style={{ whiteSpace: "pre-wrap", display: 'inline-block' }} dangerouslySetInnerHTML={{ __html: workItem }} />
                                            </Text></div>) : (<Text variant={'small'} >not present</Text>)
                                            }
                                        </div>
                                    </div>
                                </div>) : (
                                    <div>
                                        <div className="row" style={{ marginBottom: "20px" }}>
                                            <div className="col">
                                                <Text variant={'medium'} block>
                                                    <b>Tags</b>
                                                </Text>
                                                <Text>
                                                    {categories ? categories.map((t) => (
                                                        <Tag name={t} compact={false} viewState={TagViewState.TextOnly} />
                                                    )) : ""}
                                                </Text>
                                            </div>
                                        </div>
                                        <div className="row" style={{ marginBottom: "20px" }}>
                                            <div className="col">
                                                <Text variant={'medium'} block>
                                                    <b>Task</b>
                                                </Text>

                                                {workItem !== "" ? (<div><Icon iconName='TaskSolid' />&nbsp;<Text variant={'small'} >
                                                    <div style={{ whiteSpace: "pre-wrap", display: 'inline-block' }} dangerouslySetInnerHTML={{ __html: workItem }} />
                                                </Text></div>) : (<Text variant={'small'} >not present</Text>)
                                                }
                                            </div>
                                        </div>
                                    </div>)}
                            </div>

                            <Pivot
                                aria-label="Message"
                                linkFormat={'links'}
                                overflowBehavior={'menu'}
                                overflowAriaLabel="more items">

                                <PivotItem headerText="Details">
                                    &nbsp;<br />
                                    <Text variant={'medium'} block>
                                        <Separator />
                                        <b>Message summary</b><br />
                                    </Text>
                                    <Text>
                                        <div dangerouslySetInnerHTML={{ __html: description }} />
                                    </Text>
                                    <Separator />
                                    <div>
                                        {links.map((l: any) => (<Button text={l.name} href={l.url} target="_blank" styles={{ root: { margin: "8px" } }} />))}
                                    </div>
                                </PivotItem>

                                <PivotItem headerText="Activity">
                                    &nbsp;<br />
                                    {id ? (<AuditLog itemId={id} itemType={componentName} scheme='item' ref={this.activityLog} />) : (<Text variant='medium'>No activities available</Text>)}
                                </PivotItem>
                            </Pivot>


                            <AlertMessage title='' header='Something went wrong.' message={error} isBlocking={true} isOpen={error !== undefined} onClose={this._closeErrorDialog} />
                        </div>
                    </div>) : (
                    <div className="loadingProgress" style={{ display: id && error !== undefined ? 'none' : '' }}>
                        <br /><br />
                        <Spinner size={SpinnerSize.large} />
                    </div>)}
            </div>
        );
    }

    componentDidMount() {
        this._getMessageDetails(this.props.id);
    }

    componentDidUpdate(prevProps: { id?: string }) {
        if (prevProps.id !== this.props.id) {
            this.setState({
                initialized: false
            });

            this._getMessageDetails(this.props.id);
        }
    }

    private _getMessageDetails(id?: string) {
        const requiredRoles: string[] = ['ServiceHealthReader', 'Communication.Write.All', 'Admin'];
        var authResponse: AuthenticationResult;
        var mauCollection: IServiceInfo[] = [];
        var customActions: ICustomAction[] = [];
        var messageId = id;

        if (messageId !== undefined && messageId.trim().length > 0) {
            acquireAccessToken()
                .then((response) => {
                    var tokenClaims: any = response.account?.idTokenClaims;
                    const userRoles: any = tokenClaims?.roles;
                    var userHasRequiredRole: boolean = userRoles.some((r: string) => requiredRoles.includes(r));
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
                                fetch('/api/Roadmap/' + messageId, { headers: authResponse.idToken === "" ? {} : { 'Authorization': `Bearer ${authResponse.idToken}` } })
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
                                            /* var serviceCollection: IServiceInfo[] = [];
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
                                            } */

                                            var links: IMessageLink[] = [];
                                            
                                            if (result.moreInfoLink !== undefined && result.moreInfoLink.trim() !== '')
                                                links.push({
                                                    name: "More information",
                                                    url: result.moreInfoLink
                                                });

                                            if (result.link !== undefined && result.link.trim() !== '')
                                                links.push({
                                                    name: "Roadmap",
                                                    url: result.link
                                                });

                                            var categories;
                                            var tags;
                                            var cloudInstances;
                                            var products;
                                            var releasePhase;
                                            var platforms;

                                            if (!result.category || result.category.length <= 0) {
                                                categories = [];
                                            } else {
                                                if (result.category[0].tagName === undefined) {
                                                    categories = result.category;
                                                } else {
                                                    categories = result.category.map((t: IRoadmapTag) => t.tagName);
                                                }
                                            }

                                            if (!result.tags || result.tags.length <= 0) {
                                                tags = [];
                                            } else {
                                                if (result.tags[0].tagName === undefined) {
                                                    tags = result.tags;
                                                } else {
                                                    tags = result.tags.map((t: IRoadmapTag) => t.tagName);
                                                }
                                            }

                                            if (!result.cloudInstances || result.cloudInstances.length <= 0) {
                                                cloudInstances = [];
                                            } else {
                                                if (result.cloudInstances[0].tagName === undefined) {
                                                    cloudInstances = result.cloudInstances;
                                                } else {
                                                    cloudInstances = result.cloudInstances.map((t: IRoadmapTag) => t.tagName);
                                                }
                                            }

                                            if (!result.products || result.products.length <= 0) {
                                                products = [];
                                            } else {
                                                if (result.products[0].tagName === undefined) {
                                                    products = result.products;
                                                } else {
                                                    products = result.products.map((t: IRoadmapTag) => t.tagName);
                                                }
                                            }

                                            if (!result.releasePhase || result.releasePhase.length <= 0) {
                                                releasePhase = [];
                                            } else {
                                                if (result.releasePhase[0].tagName === undefined) {
                                                    releasePhase = result.releasePhase;
                                                } else {
                                                    releasePhase = result.releasePhase.map((t: IRoadmapTag) => t.tagName);
                                                }
                                            }

                                            if (!result.platforms || result.platforms.length <= 0) {
                                                platforms = [];
                                            } else {
                                                if (result.platforms[0].tagName === undefined) {
                                                    platforms = result.platforms;
                                                } else {
                                                    platforms = result.platforms.map((t: IRoadmapTag) => t.tagName);
                                                }
                                            }

                                            this.setState({
                                                id: result.id,
                                                title: result.title,
                                                links: links,
                                                status: result.status ? result.status : 'not specified',
                                                availabilityDate: result.availabilityDate,
                                                availabilityFrom: new Date(result.availabilityFrom),
                                                availabilityTo: new Date(result.availabilityTo),
                                                publicPreviewDate: result.publicPreviewDate,
                                                publicPreviewFrom: new Date(result.publicPreviewFrom),
                                                publicPreviewTo: new Date(result.publicPreviewTo),
                                                categories: categories,
                                                tags: tags,
                                                cloudInstances: cloudInstances,
                                                products: products,
                                                releasePhase: releasePhase,
                                                platforms: platforms,
                                                description: result.description,
                                                published: new Date(result.published),
                                                lastUpdated: new Date(result.lastUpdated),
                                                extendedProperties: result.extendedProperties,
                                                workItem: result.workItemID ? "<a href='" + result.workItemURL + "' target='_blank'>" + result.workItemID + "</a>" : "",
                                                initialized: true
                                            });

                                            if (this.props.onUpdateParentTitle)
                                                this.props.onUpdateParentTitle(result.title);
                                        }
                                    }));
                }).catch((err) => {
                    this.setState({
                        error: err.message
                    });
                });
        }
    }

    /* private _procecssPublishing = (): void => {
        var params: RequestInit;
        acquireAccessToken()
            .then((response) => {
                const body = {
                    id: this.state.id,
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

                fetch("/api/PublicMessage/Publish", params)
                    .then((response) => {

                        if (!response.ok) {
                            // make the promise be rejected if we didn't get a 2xx response
                            const err = new Error("Couldn't publish message. Error details:<br/><br/><b>HTTP " + response.status + "</b><br/>" + response.statusText);
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
                    id: this.state.id
                };

                params = {
                    headers: {
                        "Content-Type": "application/json charset=UTF-8",
                        "Authorization": "Bearer " + response.idToken
                    },
                    body: JSON.stringify(body),
                    method: "POST"
                };

                fetch("/api/PublicMessage/Unpublish", params)
                    .then((response) => {
                        if (!response.ok) {
                            // make the promise be rejected if we didn't get a 2xx response
                            const err = new Error("Couldn't unpublish message. Error details:<br/><br/><b>HTTP " + response.status + "</b><br/>" + response.statusText);
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

    private _onPublicCommentsChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
        this.setState({
            newComments: newValue ? newValue : ""
        });
    } */

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
                                msg += " Tried to " + (archive ? 'archive' : 'restore') + " already " + (archive ? 'archive' : 'restore') + "d message.";
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

    private _closeErrorDialog = (): void => {
        this.setState({
            error: undefined
        });
    }
}