import * as React from 'react';
import {
    SelectionMode, ICommandBarProps,
    Link, ILinkStyleProps, ILinkStyles
} from '@fluentui/react';
import { DetailPageHeader } from '@m365-admin/detail-page';
import { IconAlert, IconAlertSize, IconAlertStatus } from '@m365-admin/icon-alert';
import { CompositeList, ICompositeListColumn } from '@m365-admin/composite-list';
import { EmptyStateImageSize } from '@m365-admin/empty-state';
import { M365Breadcrumb } from '@m365-admin/m365-breadcrumb';
import { AccessDenied } from "../../AccessDenied";
import { acquireAccessToken } from "../../../auth/AccessTokenHelper";
import { AuthenticationResult } from '@azure/msal-browser';
import { IComponent } from '../notifications/Routing';
import { AdminLayout } from '../AdminLayout';

interface IComponentState {
    components: IComponent[];
    initialized: boolean;
    accessGranted: boolean;
    error?: string;
}

export class Components extends React.Component<{}, IComponentState> {

    constructor(props: {}) {
        super(props);

        this.state = {
            components: [],
            initialized: false,
            accessGranted: false,
            error: undefined
        };
    }

    public render() {
        const {
            components, initialized
        } = this.state;

        if (!initialized)
            return (<div />);

        const pipeFabricStyles = (p: ILinkStyleProps): ILinkStyles => ({
            root: {
                textDecoration: 'none',
                color: p.theme.semanticColors.bodyText
            },
        });

        const columns: ICompositeListColumn[] = [
            {
                key: 'entityName',
                name: 'Name',
                minWidth: 150,
                maxWidth: 350,
                fieldName: 'name',
                isResizable: true,
                isRowHeader: true,
                onRender: (item: IComponent) => {
                    return <Link href={'/admin/customActions/actions?entityId=' + item.id} styles={pipeFabricStyles}>{item.name}</Link >;
                }
            }
        ];

        const commandBarProps: ICommandBarProps = {
            items: []
        };

        var listData = [
            {
                listProps: {
                    checkButtonAriaLabel: 'Select item',
                    ariaLabelForSelectAllCheckbox: 'Select all items',
                    items: components,
                    columns: columns,
                    selectionMode: SelectionMode.none,
                    ariaLabelForGrid: 'Component list, use arrows to navigate'
                },
                key: 'connectors'
            }
        ];

        return (
            <AdminLayout>
                <div className="container" style={{ maxWidth: '97%' }}>
                    <div className="row">
                        <div className="col">
                            <M365Breadcrumb
                                items={[
                                    { text: 'Home', key: 'home', href: '/admin' },
                                    { text: 'Integration', key: 'customActions', href: '/admin/customActions' },
                                    { text: 'Components', key: 'components', isCurrentItem: true }
                                ]}
                                style={{ marginBottom: '16px' }}
                            />
                            <DetailPageHeader
                                title="Integration"
                                description="Select a component from the list below to configure custom actions."
                            />
                        </div>
                    </div>
                    <div className="row">
                        <div className="col">
                            <CompositeList
                                detailsListProps={listData}
                                commandBarProps={commandBarProps}
                                defaultIsCompactMode={true}
                                onSelectionChanged={this._doNothing}
                                legacyInvokeAndSelect={true}
                                isEmpty={
                                    listData === undefined ||
                                    listData[0].listProps?.items === undefined ||
                                    listData[0].listProps?.items?.length === 0}
                                emptyStateProps={{
                                    title: 'No components found',
                                    body: 'Please contact your administrator.',
                                    imageProps: {
                                        src: '/images/empty.png',
                                        alt: 'No components found graphic'
                                    },
                                    illustrationSize: EmptyStateImageSize.Large
                                }}
                            />
                        </div>
                    </div>
                </div>
            </AdminLayout>
        );
    }

    componentDidMount() {
        this._onLoadConfig();
    }

    private _onLoadConfig(): void {
        const requiredRoles: string[] = ['Admin'];
        var authResponse: AuthenticationResult;
        var userHasRequiredRole: boolean = false;
        var componentCollection: IComponent[] = [];

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
                    fetch('/api/Components', { headers: authResponse.idToken === "" ? {} : { 'Authorization': `Bearer ${authResponse.idToken}` } })
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
                            for (const component of result) {
                                if (component.capabilities?.includes('CustomActions'))
                                    componentCollection.push({
                                        key: component.id,
                                        id: component.id,
                                        name: component.name,
                                        internalName: component.internalName,
                                        icon: component.icon,
                                        capabilities: component.capabilities,
                                        entityProperties: component.entityProperties
                                    });
                            }

                            componentCollection = _copyAndSort(componentCollection, 'name', false);

                            this.setState({
                                components: componentCollection,
                                initialized: true
                            });
                        }).catch((err) => {
                            this.setState({
                                error: err.message
                            });
                        });
            });
    }

    private _doNothing = (): void => {

    }
}

function _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
    const key = columnKey as keyof T;
    return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
}