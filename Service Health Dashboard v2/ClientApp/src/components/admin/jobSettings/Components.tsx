import * as React from 'react';
import {
    SelectionMode, IObjectWithKey, ICommandBarProps,
    Panel, PanelType, Dialog, DialogType, TextField, Link, ILinkStyleProps, ILinkStyles,
    PrimaryButton, DefaultButton, DialogFooter, Label, Dropdown, IDropdownOption, ISelectableOption
} from '@fluentui/react';
import { DetailPageHeader } from '@m365-admin/detail-page';
import { IconAlert, IconAlertSize, IconAlertStatus } from '@m365-admin/icon-alert';
import { CompositeList, ICompositeListColumn, ICompositeListSelectionMap, ICompositeListSelectionMapItem } from '@m365-admin/composite-list';
import { EmptyStateImageSize } from '@m365-admin/empty-state';
import { M365Breadcrumb } from '@m365-admin/m365-breadcrumb';
import { Tag, TagType } from '@m365-admin/tag';
import { AccessDenied } from "../../AccessDenied";
import { acquireAccessToken } from "../../../auth/AccessTokenHelper";
import { AuthenticationResult } from '@azure/msal-browser';
import { AdminLayout } from '../AdminLayout';

interface IJobInfo {
    icon?: string;
    id: string;
    name: string;
    lastRun?: Date;
    duration?: string;
    state: string;
}
interface IJobComponentsState {
    components: IJobInfo[];
    initialized: boolean;
    accessGranted: boolean;
    error?: string;
}

export class JobComponents extends React.Component<{}, IJobComponentsState> {

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
                onRender: (item: IJobInfo) => {
                    return <Link href={'/admin/jobs/settings?id=' + item.id} styles={pipeFabricStyles}>{item.name}</Link >;
                }
            },
            {
                key: 'lastRuntime',
                name: 'Last run',
                minWidth: 128,
                maxWidth: 128,
                fieldName: 'lastRun',
                isResizable: false,
                onRender: (item: IJobInfo) => {
                    return item.lastRun ? new Date(item.lastRun).toLocaleString() : ""
                }
            },
            {
                key: 'lastDuration',
                name: 'Duration',
                minWidth: 128,
                maxWidth: 128,
                fieldName: 'duration',
                isResizable: false
            },
            {
                key: 'state',
                name: 'State',
                minWidth: 64,
                maxWidth: 64,
                isResizable: false,
                onRender: (item: IJobInfo) => {
                    if (item.state !== null && item.state.trim() === "")
                        return (<></>)
                    else {
                        var tagType: TagType = TagType.Insight;

                        switch (item.state?.toLowerCase()) {
                            case 'completed':
                                tagType = TagType.Insight;
                                break;
                            case 'failed':
                                tagType = TagType.ActiveWarning;
                                break;
                            case 'running':
                                tagType = TagType.HighImpactInformational;
                                break;
                            default:
                                tagType = TagType.LowImpactInformational;
                        }

                        return item.state !== null ? (<Tag key={'tag-' + item.id} tagType={tagType}>{item.state}</Tag>) : ("");
                    }
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
                    ariaLabelForGrid: 'Job list, use arrows to navigate'
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
                                    { text: 'Jobs', key: 'jobs', isCurrentItem: true }
                                ]}
                                style={{ marginBottom: '16px' }}
                            />
                            <DetailPageHeader
                                title="Jobs"
                                description="Select a component from the list below to configure job settings and view statistics."
                            />
                        </div>
                    </div>
                    <div className="row">
                        <div className="col">
                            <CompositeList
                                detailsListProps={listData}
                                commandBarProps={commandBarProps}
                                defaultIsCompactMode={false}
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
        var componentCollection: IJobInfo[] = [];

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
                    fetch('/api/jobs', { headers: authResponse.idToken === "" ? {} : { 'Authorization': `Bearer ${authResponse.idToken}` } })
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
                            componentCollection = _copyAndSort(result, 'name', false);
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