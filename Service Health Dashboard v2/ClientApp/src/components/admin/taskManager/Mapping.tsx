import * as React from 'react';
import { Text, Pivot, PivotItem } from '@fluentui/react';
import { M365Breadcrumb } from '@m365-admin/m365-breadcrumb';
import { DetailPageHeader } from '@m365-admin/detail-page';
import { AdminLayout } from "../AdminLayout";
import { ConfigEditor } from "../jsonConfigEditor/ConfigEditor";
import { IComponent } from '../notifications/Routing';
import { acquireAccessToken } from "../../../auth/AccessTokenHelper";
import { AuthenticationResult } from '@azure/msal-browser';

interface IAdminMappingState {
    entityId: string | undefined,
    component: IComponent | undefined;
    error: string | undefined;
    initialized: boolean;
    accessGranted: boolean;
}

export class AdminTaskMapping extends React.Component<{}, IAdminMappingState> {
    constructor(props: {}) {
        super(props);

        const queryParams = new URLSearchParams(window.location.search);
        const entityId = queryParams.get('entityId');

        this.state = {
            entityId: entityId ? entityId : undefined,
            error: undefined,
            component: undefined,
            initialized: false,
            accessGranted: false
        };
    }

    public render() {
        const {
            component
        } = this.state;

        return (
            <AdminLayout>
                <div className="container" style={{ maxWidth: '97%' }}>
                    <div className="row">
                        <div className="col">
                            <M365Breadcrumb
                                items={[
                                    { text: 'Home', key: 'home', href: '/admin' },
                                    { text: 'Mapping', key: 'notifications' },
                                    { text: 'Components', key: 'components', href: '/admin/mapping' },
                                    { text: component ? component.name : 'Undefined', key: component ? component.internalName : 'undefined', isCurrentItem: true }
                                ]}
                                style={{ marginBottom: '16px' }}
                            />
                            <DetailPageHeader
                                title="Task field mapping"
                                description="Use editor below to define mapping between Service Health Hub metadata and task fields in target system."
                            />
                        </div>
                    </div>
                    <div className="row">
                        <div className="col">
                            {component ? (
                                <ConfigEditor component={component.internalName} element='metadataMapping' />
                                ) : ("")}
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
        var component: IComponent | undefined = undefined;

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
            }).then(() => {
                if (userHasRequiredRole)
                    fetch('/api/Components?id=' + this.state.entityId, { headers: authResponse.idToken === "" ? {} : { 'Authorization': `Bearer ${authResponse.idToken}` } })
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
                            if (result !== undefined && result.length > 0)
                                component = result[0] as IComponent;

                            this.setState({
                                component: component
                            });
                        }).catch((err) => {
                            this.setState({
                                error: err.message
                            });
                        });
            });
    }
}