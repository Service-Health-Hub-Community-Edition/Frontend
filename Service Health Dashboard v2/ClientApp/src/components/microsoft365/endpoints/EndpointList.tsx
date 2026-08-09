import * as React from 'react';
import type {
    IButtonStyles,
    IColumn,
    ICommandBarItemProps,
    ICommandBarProps,
    IDetailsRowProps,
    IDropdownOption,
    IGroup,
    IIconProps,
    IRenderFunction,
    IStyle,
    ITheme
} from '@fluentui/react';
import {
    Announced, FontWeights, FontIcon, SelectionMode, Text,
    Link, ILinkStyles, ILinkStyleProps, ThemeProvider,
    mergeStyles, mergeStyleSets, TooltipHost, Tooltip, Panel, PanelType,
    PrimaryButton, Pivot, PivotItem, IBreadcrumbItem
} from '@fluentui/react';
import type {
    ICompositeListColumn,
    ICompositeListDetailsList,
    ICompositeListFarItemOptions,
    ICompositeListSelectionMap,
    ICompositeListSelectionMapItem,
    IListSettingsColumn,
    IListSettingsColumnGroup,
} from '@m365-admin/composite-list';
import { ListSettings } from '@m365-admin/composite-list';
import { CompositeList, CompositeListDetailsList, CompositeListRow } from '@m365-admin/composite-list';
import { MetaDataList, MetaDataItem } from '@m365-admin/metadata';
import { EmptyStateImageSize } from '@m365-admin/empty-state';
import { LoadingPane, LoadingAnimationType } from '@m365-admin/loading';
import { Collapsible } from '@m365-admin/collapsible';
import { M365Breadcrumb } from '@m365-admin/m365-breadcrumb';
import { DetailPageHeader } from '@m365-admin/detail-page';
import { GlobalState } from '../../GlobalState';
import { acquireAccessToken } from "../../../auth/AccessTokenHelper";

export interface IEndpointSet {
    id: number;
    serviceArea: string;
    serviceAreaDisplayName: string;
    category: string;
    urls: string[] | undefined;
    ips: string[] | undefined;
    tcpPorts: string | undefined;
    udpPorts: string | undefined;
    expressRoute: boolean;
    required: boolean;
    notes: string | undefined;
    changes: any | undefined;
}

interface IEndpointSetsHeaderItem {
    name: string;
}

export interface IEndpointListState {
    items: IEndpointSet[];
    groupNames: string[];
    groups: IGroup[];
    selectedItem: IEndpointSet | undefined;
    isDetailsPanelOpen: boolean;
    accessGranted: boolean;
    error: string | undefined;
    initialized: boolean;
    loadingText: string;
    theme?: any;
}

// Yes: CompletedSolid, No: StatusErrorFull
// Optimize: GroupList - NumberedList - BulletedTreeList, Allow: CheckList, Default: AppIconDefaultList,

const iconClass = mergeStyles({
    fontSize: 16,
    height: 16,
    width: 16,
    margin: "0 16px"
});

const iconClass8 = mergeStyles({
    fontSize: 16,
    height: 16,
    width: 16,
    margin: "0 4px"
});

const classNames = mergeStyleSets({
    error: [{ color: '#dc3545' }, iconClass],
    ok: [{ color: '#107c10' }, iconClass],
    information: [{ color: 'dodgerblue' }, iconClass],
    warning: [{ color: '#ffb900' }, iconClass],
    sync: [{ color: 'dodgerblue' }, iconClass]
});

const classNames8 = mergeStyleSets({
    error: [{ color: '#dc3545' }, iconClass8],
    ok: [{ color: '#107c10' }, iconClass8],
    information: [{ color: 'dodgerblue' }, iconClass8],
    warning: [{ color: '#ffb900' }, iconClass8],
    sync: [{ color: 'dodgerblue' }, iconClass8]
});

const iconList: any = {
    Failed: 'StatusCircleErrorX',
    Running: 'SyncStatusSolid',
    Success: 'StatusCircleCheckmark',
    Warning: 'WarningSolid',
    Information: 'Info',
    Optimize: 'BulletedTreeList',
    Allow: 'CheckList',
    Default: 'AppIconDefaultList'
}



const headerColumns: IColumn[] = [
    {
        key: 'serviceAreaDisplayName',
        name: '',
        fieldName: 'serviceAreaDisplayName',
        isResizable: true,
        minWidth: 100,
        maxWidth: 200,
        onRender: (item: IEndpointSet) => {
            return (
                <Text styles={{ root: { fontWeight: FontWeights.semibold } }}>
                    {item.serviceAreaDisplayName}
                </Text>
            );
        },
    }
];

interface IEndpointSetChangeState {
    change: any;
    isOpen: boolean;
    title: string;
}

class EndpointSetChange extends React.Component<{ change: any, isOpen?: boolean }, IEndpointSetChangeState> {
    constructor(props: { change: any, isOpen?: boolean }) {
        super(props);

        this.state = {
            isOpen: props.isOpen ? props.isOpen : false,
            change: props.change,
            title: props.change?.version ?
                new Date(props.change?.version.substring(0, 4) + '-' + props.change?.version.substring(4, 6) + '-' + props.change?.version.substring(6, 8))
                    .toLocaleDateString([], { year: 'numeric', month: 'short', day: 'numeric' }) : 'Not specified'
        }
    }

    render() {
        const { title, change, isOpen } = this.state;

        return (
            <Collapsible
                title={title + ': ' + change?.disposition + (change.impact && change?.impact !== null ? ' (' + change?.impact + ')' : '')}
                key={_createGuid()}>
                <MetaDataList>
                    <MetaDataItem header='Id' body={change?.id} />
                    <MetaDataItem header='Disposition' body={change?.disposition} />
                    <MetaDataItem header='Impact' body={change?.impact} />
                    <MetaDataItem header='Version' body={change?.version} />
                </MetaDataList>

                {change?.remove ?
                    (<div>
                        <br />
                        <Text variant='medium' style={{ marginBottom: '16px' }} >
                            <b>Removed</b>
                            <br />&nbsp;
                        </Text>
                        
                        <MetaDataList styles={{ itemsContainer: { backgroundColor: '#f7f7f7' } }} >
                            <MetaDataItem header='URLs' body={change?.remove?.urls?.map((i: any) => (<div>{i}<br /></div>))} />
                            <MetaDataItem header='IP addresses and ranges' body={change?.remove?.ips?.map((i: any) => (<div>{i}<br /></div>))} />
                        </MetaDataList>
                    </div>
                        ) : ''}

                {change?.add ?
                    (<div>
                        <br />
                        <br />
                        <Text variant='medium' style={{ marginBottom: '16px' }} >
                            <b>Added</b>
                            <br />&nbsp;
                        </Text>

                        <MetaDataList styles={{ itemsContainer: { backgroundColor: '#f7f7f7' } }} >
                            <MetaDataItem header='Effective date' body={change?.add?.effectiveDate ? 
                                new Date(change?.add?.effectiveDate.substring(0, 4) + '-' + change?.add?.effectiveDate.substring(4, 6) + '-' + change?.add?.effectiveDate.substring(6, 8))
                                    .toLocaleDateString([], { year: 'numeric', month: 'short', day: 'numeric' }) : ''
                                } />
                        </MetaDataList>
                        <MetaDataList styles={{ itemsContainer: { backgroundColor: '#f7f7f7' } }}>
                            <MetaDataItem header='URLs' body={change?.add?.urls?.map((i: any) => (<div>{i}<br /></div>))} />
                            <MetaDataItem header='IP addresses and ranges' body={change?.add?.ips?.map((i: any) => (<div>{i}<br /></div>))} />
                        </MetaDataList>
                    </div>
                    ) : ''}

                {change?.previous ?
                    (<div>
                        <br />
                        <Text variant='medium' style={{ marginBottom: '16px' }} >
                            <b>Previous</b>
                            <br />&nbsp;
                        </Text>

                        <MetaDataList styles={{ itemsContainer: { backgroundColor: '#f7f7f7' } }} >
                            {change?.previous?.serviceArea || change?.previous?.serviceArea === null ? (
                                <MetaDataItem header='Service area' body={change?.previous?.serviceArea} />
                            ) : ''}
                            {change?.previous?.category || change?.previous?.category === null ? (
                                <MetaDataItem header='Category' body={change?.previous?.category} />
                            ) : ''}
                            {change?.previous?.required || change?.previous?.required === null ? (
                                <MetaDataItem header='Required' body={change?.previous?.required ? 'Yes' : 'No'} />
                            ) : ''}
                            {change?.previous?.expressRoute || change?.previous?.expressRoute === null ? (
                                <MetaDataItem header='Express route' body={change?.previous?.expressRoute ? 'Yes' : 'No'} />
                            ) : ''}
                            {change?.previous?.notes || change?.previous?.notes === null ? (
                                <MetaDataItem header='Notes' body={change?.previous?.notes} />
                            ) : ''}
                            {change?.previous?.ipPorts || change?.previous?.ipPorts === null ? (
                                <MetaDataItem header='IP ports' body={change?.previous?.ipPorts } />
                            ) : ''}
                            {change?.previous?.udpPorts || change?.previous?.udpPorts === null ? (
                                <MetaDataItem header='UDP ports' body={change?.previous?.udpPorts } />
                            ) : ''}

                        </MetaDataList>
                    </div>
                    ) : ''}

                {change?.current ?
                    (<div>
                        <br />
                        <Text variant='medium' style={{ marginBottom: '16px' }} >
                            <b>Current</b>
                            <br />&nbsp;
                        </Text>

                        <MetaDataList styles={{ itemsContainer: { backgroundColor: '#f7f7f7' } }} >
                            {change?.current?.serviceArea || change?.current?.serviceArea === null ? (
                                <MetaDataItem header='Service area' body={change?.current?.serviceArea} />
                            ) : ''}
                            {change?.current?.category || change?.current?.category === null ? (
                                <MetaDataItem header='Category' body={change?.current?.category} />
                            ) : ''}
                            {change?.current?.required || change?.current?.required === null ? (
                                <MetaDataItem header='Required' body={change?.current?.required ? 'Yes' : 'No'} />
                            ) : ''}
                            {change?.current?.expressRoute || change?.current?.expressRoute === null ? (
                                <MetaDataItem header='Express route' body={change?.current?.expressRoute ? 'Yes' : 'No'} />
                            ) : ''}
                            {change?.current?.notes || change?.current?.notes === null ? (
                                <MetaDataItem header='Notes' body={change?.current?.notes} />
                            ) : ''}
                            {change?.current?.ipPorts || change?.current?.ipPorts === null ? (
                                <MetaDataItem header='IP ports' body={change?.current?.ipPorts } />
                            ) : ''}
                            {change?.current?.udpPorts || change?.current?.udpPorts === null ? (
                                <MetaDataItem header='UDP ports' body={change?.current?.udpPorts } />
                            ) : ''}

                        </MetaDataList>
                    </div>
                    ) : ''}
                
            </Collapsible>
            );
    }
}

const instance: string = 'Worldwide';

export class EndpointList extends React.Component<{ message?: string }, IEndpointListState> {
    static contextType = GlobalState;

    pipeFabricStyles = (p: ILinkStyleProps): ILinkStyles => ({
        root: {
            textDecoration: 'none',
            color: p.theme.semanticColors.bodyText
        },
    });

    columns: ICompositeListColumn[] = [
        {
            key: 'connectivity',
            name: 'Connectivity',
            minWidth: 125,
            maxWidth: 125,
            fieldName: 'category',
            isResizable: true,
            isRowHeader: true,
            onRender: (item: IEndpointSet) => {
                var iconName: string = '';
                var iconClass: any = undefined;
                switch (item.category.toLowerCase()) {
                    case 'optimize':
                        iconName = iconList.Optimize;
                        iconClass = classNames.information;
                        break;

                    case 'allow':
                        iconName = iconList.Allow;
                        iconClass = classNames.ok;
                        break;

                    default:
                        iconName = iconList.Default;
                        iconClass = classNames.information;
                }

                return <>

                    <FontIcon aria-label={item.category} iconName={iconName} className={iconClass} />
                    <Link onClick={(event) => {
                        event.preventDefault();
                        this._openDetails(item.id);
                    }} styles={this.pipeFabricStyles}>{item.category}</Link >
                    </>;
            },
        },
        {
            key: 'changes',
            name: 'Changes',
            minWidth: 50,
            maxWidth: 50,
            fieldName: 'changes',
            isResizable: true,
            onRender: (item: IEndpointSet) => {
                return <>{item.changes && item.changes.length > 0 ? item.changes.length : '---'}</>;
            },
        },
        {
            key: 'id',
            name: 'Id',
            minWidth: 50,
            maxWidth: 50,
            fieldName: 'id',
            isResizable: false
        },
        {
            key: 'urls',
            name: 'URLs',
            minWidth: 100,
            maxWidth: 200,
            fieldName: 'urls',
            isResizable: true,
            onRender: (item: IEndpointSet) => {
                return <>{item.urls ? item.urls.join(', ') : ''}</>;
            }
        },
        {
            key: 'ips',
            name: 'IPs',
            minWidth: 100,
            maxWidth: 200,
            fieldName: 'ips',
            isResizable: true,
            onRender: (item: IEndpointSet) => {
                return <>{item.ips ? item.ips.join(', ') : ''}</>;
            }
        },
        {
            key: 'tcpPorts',
            name: 'TCP',
            minWidth: 50,
            maxWidth: 50,
            fieldName: 'tcpPorts',
            isResizable: false
        },
        {
            key: 'udpPorts',
            name: 'UDP',
            minWidth: 50,
            maxWidth: 50,
            fieldName: 'udpPorts',
            isResizable: false
        },
        {
            key: 'expressRoute',
            name: 'Express route',
            minWidth: 75,
            maxWidth: 75,
            fieldName: 'expressRoute',
            isResizable: false,
            onRender: (item: IEndpointSet) => {
                return <>{item.expressRoute ?
                    (<FontIcon aria-label={'Yes'} iconName={iconList.Success} className={classNames.ok} style={{ cursor: 'default' }} />) :
                    (<FontIcon aria-label={'No'} iconName={iconList.Failed} className={classNames.error} style={{ cursor: 'default' }} />)}</>;
            }
        },
        {
            key: 'required',
            name: 'Required',
            minWidth: 75,
            maxWidth: 150,
            fieldName: 'required',
            isResizable: false,
            onRender: (item: IEndpointSet) => {
                return <>
                    {item.required ?
                        (<FontIcon aria-label={'Yes'} iconName={iconList.Success} className={classNames8.ok} style={{ cursor: 'default' }} />) :
                        (<FontIcon aria-label={'No'} iconName={iconList.Failed} className={classNames8.error} style={{ cursor: 'default' }} />)}
                    {item.notes ?
                        (<TooltipHost content={item.notes} >
                            <FontIcon aria-label={'Notes'} iconName={iconList.Information} className={classNames8.information} style={{ cursor: 'default' }} />
                        </TooltipHost>) : ''}
                </>;
            }
        }
    ];

    _breadcrumbItems: IBreadcrumbItem[] = [
        { text: 'Home', key: 'home', isCurrentItem: false, href: '/' },
        { text: 'Office 365 Endpoints', key: 'o365EndpointSets', isCurrentItem: true }
    ];

    constructor(props: { message?: string }) {
        super(props);

        this.state = {
            items: [],
            groupNames: [],
            groups: [],
            selectedItem: undefined,
            isDetailsPanelOpen: false,
            accessGranted: false,
            error: undefined,
            initialized: false,
            loadingText: ''
        };
    }

    public render() {
        const { items, groups, selectedItem, isDetailsPanelOpen, accessGranted, error, initialized, loadingText, theme } = this.state;

        if (theme === undefined)
            return "";

        return (
            <div style={{ overflowX: 'hidden' }}>
                    <M365Breadcrumb
                        items={this._breadcrumbItems}
                        style={{ marginBottom: '16px' }}
                        ariaLabel="Breadcrumb with items rendered as buttons"
                        overflowAriaLabel="More links"
                    />
                    <DetailPageHeader
                        title="Office 365 Endpoints"
                        description="The Office 365 IP Address and URL list helps you better identify and differentiate Office 365 network traffic, making it easier for you to evaluate, configure, and stay up to date with changes."
                    />
                    <CompositeList
                        defaultIsCompactMode={true}
                        commandBarProps={{
                            items: []
                        }}
                        detailsListProps={[
                            {
                                listProps: {
                                    columns: this.columns,
                                    items: items,
                                    groups: groups,
                                    selectionMode: SelectionMode.none
                                },
                                key: 'endpointSets'
                            }
                        ]}
                    />
                    {!initialized ? (
                        <LoadingPane
                            isLoading={!initialized}
                            loadingAnimationType={LoadingAnimationType.FluentSpinner}
                            animationAriaLabel="Loading data"
                            primaryLoadingText={'Hold on, the data is on its way'}
                            secondaryLoadingText={loadingText}
                        />
                    ): '' }
                    <Panel
                        isOpen={isDetailsPanelOpen}
                        isBlocking={true}
                        isLightDismiss={true}
                        onDismiss={() => this._dismissDetails()}
                        onDismissed={() => this._onDetailsDismissed()}
                        type={PanelType.medium}                      
                        headerText={selectedItem?.serviceAreaDisplayName + ' (' + selectedItem?.id + ')'}
                    >
                        <Pivot
                            aria-label="Endpoint set"
                            linkFormat={'links'}
                            overflowBehavior={'menu'}
                            overflowAriaLabel="more items">

                            <PivotItem headerText="Details">
                                &nbsp;<br />
                                <MetaDataList>
                                    <MetaDataItem header='Category' body={selectedItem?.category} />
                                    <MetaDataItem header='Required' body={selectedItem?.required ? 'Yes' : 'No'} />
                                    <MetaDataItem header='Express route' body={selectedItem?.expressRoute ? 'Yes' : 'No'} />
                                    {selectedItem?.notes ? (<MetaDataItem header='Notes' body={selectedItem?.notes} />) : ''}
                                </MetaDataList>
                                <br />
                                <MetaDataList>
                                    <MetaDataItem header='URLs' body={selectedItem?.urls?.map((i) => (<div>{i}<br /></div>))} />
                                    <MetaDataItem header='IP addresses and ranges' body={selectedItem?.ips?.map((i) => (<div>{i}<br /></div>))} />
                                    {selectedItem?.tcpPorts ? (<MetaDataItem header='TCP ports' body={selectedItem?.tcpPorts} />) : ''}
                                    {selectedItem?.udpPorts ? (<MetaDataItem header='UDP ports' body={selectedItem?.udpPorts} />) : ''}
                                </MetaDataList>
                            </PivotItem>

                            {selectedItem?.changes?.length > 0 ? (
                                <PivotItem headerText={"Changes (" + selectedItem?.changes?.length + ")"}>
                                    &nbsp;<br />
                                    {selectedItem?.changes?.map((c: any) => (<EndpointSetChange change={c} />))}
                                </PivotItem>) : ''}
                        </Pivot>               
                        
                        <br />
                        
                        <br/>
                        <PrimaryButton text='Close' onClick={() => this._dismissDetails()} />
                    </Panel>
            </div>
        );
    }

    componentDidMount() {
        let globalState: any = this.context;
        var theme: ITheme = globalState.getTheme();

        this.setState({
            theme: theme
        });

        this._loadData();
    }

    private _loadData(): void {
        var items: IEndpointSet[] = [];
        var groups: string[] = [];
        const requiredRoles: string[] = ['ServiceHealthReader', 'Communication.Write.All', 'Admin'];

        this.setState({
            initialized: false,
            loadingText: ''
        });

        setTimeout(() => {
            this.setState({
                loadingText: 'This is taking a little longer than expected'
            });
        }, 2000);

        acquireAccessToken()
            .then((response) => {
                var tokenClaims: any = response.account?.idTokenClaims;
                const userRoles: any = tokenClaims?.roles;
                var userHasRequiredRole: boolean = userRoles.some((r: string) => requiredRoles.includes(r));

                this.setState({
                    accessGranted: userHasRequiredRole
                });

                if (userHasRequiredRole)
                    fetch('/api/EndpointSets', { headers: response.idToken === "" ? {} : { 'Authorization': `Bearer ${response.idToken}` } })
                        .then(response => {
                            if (response.ok) {
                                return response.json();
                            } else {
                                throw Error(response.status + " " + response.statusText);
                            }
                        })
                        .then(result => {
                            for (const item of result) {
                                items.push({
                                    id: item.id,
                                    serviceArea: item.serviceArea,
                                    serviceAreaDisplayName: item.serviceAreaDisplayName,
                                    category: item.category,
                                    changes: item.changes ? item.changes.sort((a: any, b: any) => a.version > b.version ? -1 : 1) : undefined,
                                    urls: item.urls,
                                    ips: item.ips,
                                    tcpPorts: item.tcpPorts,
                                    udpPorts: item.udpPorts,
                                    expressRoute: item.expressRoute,
                                    required: item.required,
                                    notes: item.notes
                                });
                            }
                        }).then(() => {                           
                            items = items.sort((a, b) => a.id > b.id ? -1 : 1).sort((a, b) => a.serviceAreaDisplayName > b.serviceAreaDisplayName ? 1 : -1);
                            groups = Array.from(new Set(items.map(i => i.serviceAreaDisplayName)));
                            this.setState({
                                items: items,
                                groupNames: groups,
                                initialized: true
                            });

                            this.setState({
                                groups: this._getGroups(items)
                            });
                        })
                        .catch(error => {
                            this.setState({
                                error: error
                            });
                        });
            });
    }

    private _getGroups (groupItems: IEndpointSet[]): IGroup[] {
        const serviceGroups: IGroup[] = [];
        let groupsStartIndex = 0;

        this.state.groupNames.forEach((serviceAreaName: string) => {
            const endpointSets = groupItems.filter((set: IEndpointSet) => {
                return set.serviceAreaDisplayName.toLowerCase() === serviceAreaName.toLowerCase();
            });

            const setGroupCount = endpointSets.length;
            const headerItem: IEndpointSetsHeaderItem = {
                name: serviceAreaName
            };

            const group: IGroup = {
                key: serviceAreaName,
                name: serviceAreaName,
                startIndex: groupsStartIndex,
                count: setGroupCount,
                data: headerItem,
            };

            groupsStartIndex += setGroupCount;

            serviceGroups.push(group);
        });

        return serviceGroups;
    };

    private _openDetails(id: number): void {
        var selectedItem: IEndpointSet | undefined = this.state.items.find((i) => i.id === id);
        this.setState({
            selectedItem: selectedItem,
            isDetailsPanelOpen: selectedItem !== undefined
        });
    }

    private _dismissDetails(): void {
        this.setState({
            isDetailsPanelOpen: false
        });
    }

    private _onDetailsDismissed(): void {
        this.setState({
            selectedItem: undefined
        });
    }
}

function _createGuid(): string {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
        var r = Math.random() * 16 | 0, v = c === 'x' ? r : (r & 0x3 | 0x8);
        return v.toString(16);
    });
}