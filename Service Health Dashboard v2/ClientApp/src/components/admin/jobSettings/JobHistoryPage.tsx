import * as React from 'react';
import { Text, Spinner, SpinnerSize, CommandBar, SearchBox, Toggle } from '@fluentui/react';
import { DetailPageHeader } from '@m365-admin/detail-page';
import { M365Breadcrumb } from '@m365-admin/m365-breadcrumb';
import { AccessDenied } from "../../AccessDenied";
import { acquireAccessToken } from "../../../auth/AccessTokenHelper";
import { AuthenticationResult } from '@azure/msal-browser';
import { IComponent } from '../notifications/Routing';
import { AdminLayout } from '../AdminLayout';
import { JobHistory } from './JobHistory';

interface IJobHistoryPageState {
    componentId: string;
    componentName: string;
    filter?: string;
    showAll: boolean;
    initialized: boolean;
    accessGranted: boolean;
    error?: string;
}

export class JobHistoryPage extends React.Component<{}, IJobHistoryPageState> {

    constructor(props: {}) {
        super(props);

        const queryParams = new URLSearchParams(window.location.search);
        const componentId = queryParams.get('id');

        this.state = {
            componentId: componentId && componentId !== null ? componentId : "",
            componentName: "",
            showAll: false,
            accessGranted: false,
            initialized: false
        };
    }

    public render() {
        const { componentId, componentName, filter, showAll, initialized, error, accessGranted } = this.state;

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

        return (
            <AdminLayout>
                <div className="container" style={{ maxWidth: '97%' }}>
                    <div className="row">
                        <div className="col">
                            <M365Breadcrumb
                                items={[
                                    { text: 'Home', key: 'home', href: '/admin' },
                                    { text: 'Jobs', key: 'jobs', href: '/admin/jobs' },
                                    { text: componentName, key: componentName, href:'/admin/jobs/settings?id=' + componentId },
                                    { text: 'History', key: 'history', isCurrentItem: true }
                                ]}
                                style={{ marginBottom: '16px' }}
                            />
                            <DetailPageHeader
                                title={componentName + ' history'}
                                description="Job history includes job status, number of created and updated items and tasks, sent notifications and amount of failures per job. Use search box to find jobs which processed a certain service communication. Click the time stamp to view activity log for the job. Up to 200 items are available in the job history list."
                            />
                        </div>
                    </div>
                    <div className="row">
                        <div className="col">
                            <div>
                                <CommandBar items={[]} farItems={[
                                    {
                                        key: 'hideWithoutUpdates',
                                        onRender: () => {
                                            return <div style={{ paddingLeft: 8, paddingRight: 8 }}><Toggle
                                                label='Show all'
                                                inlineLabel={true}
                                                checked={showAll}
                                                onChange={(event: React.MouseEvent<HTMLElement>, checked?: boolean) => this._onSetShowAll(checked ? checked : false)} /></div>;
                                        },
                                    },
                                    {
                                        key: 'searchBox',
                                        onRender: () => {
                                            return <div style={{ paddingLeft: 8, paddingRight: 8 }}><SearchBox placeholder="Communication ID" iconProps={{ iconName: 'Search' }} onChange={this._onChangeSearchText} /></div>;
                                        },
                                    }
                                ]} />
                            </div>
                            <JobHistory componentId={componentId} communicationId={filter} maxItems={200} hideEmpty={!showAll} />
                        </div>
                    </div>
                </div>
            </AdminLayout>
        );
    }

    componentDidMount() {
        this._loadData();
    }

    _loadData() {
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
                    fetch('/api/Components?id=' + this.state.componentId, { headers: authResponse.idToken === "" ? {} : { 'Authorization': `Bearer ${authResponse.idToken}` } })
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
                                componentName: component ? component?.name : "",
                                initialized: true
                            });
                        }).catch((err) => {
                            this.setState({
                                error: err.message,
                                initialized: true
                            });
                        });
            });
    }

    private _onChangeSearchText = (ev?: React.ChangeEvent<HTMLInputElement>, newValue?: string): void => {
        this.setState({
            filter: newValue
        });
    };

    private _onSetShowAll(checked: boolean) {
        this.setState({
            showAll: checked
        });
    }
}