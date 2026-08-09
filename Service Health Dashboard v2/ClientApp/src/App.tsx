import * as React from 'react';
import * as microsoftTeams from "@microsoft/teams-js";
import { Component } from 'react';
import { Route, Switch } from 'react-router';
import { GlobalState, IGlobalState } from './components/GlobalState';
import type { IRearrangeableGridDefinition } from '@m365-admin/rearrangeable-grid';
import { UserProfile, setPropertyAsync } from './components/UserProfile';
import { NotificationState, INotification, Notifications } from './components/Notifications';
import { Layout } from './components/Layout';
import { ServiceDetailsList } from './components/ServiceDetails';
import { MessagesList } from './components/Messages';
import { IncidentReport } from './components/IncidentReports';
import { LicenseReport } from './components/LicenseReport';
import { Roadmap } from './components/RoadmapTimeline';
import {
    SharePointOnlineReport, ExchangeOnlineReport, ODBOnlineReport,
    SFBOnlineReport, MicrosoftTeamsOnlineReport, Dynamics365Report,
    M365SuiteReport, PlannerReport
} from './components/reports/ServiceReports';
import { AccessDenied } from './components/AccessDenied';
import { Configure } from './components/Configure';
import { Settings } from './components/Settings';
import { AdminLayout } from './components/admin/AdminLayout';
import { AdminHome } from './components/admin/Home';
import { AdminRouting } from './components/admin/Routing';
import { TaskMappingEntities } from './components/admin/taskManager/Entities';
import { TagDefinitionsList } from './components/admin/applicationSettings/TagDefinitions';
import { AdminTaskMapping } from './components/admin/taskManager/Mapping';
import './custom.css'
import AuthenticationTemplate from './components/auth/AuthenticationTemplate';
import AuthStart from './components/auth/AuthStart';
import AuthEnd from './components/auth/AuthEnd';
import { NotificationsConfig } from './components/admin/notifications/NotificationsConfig';
import { NotificationsRouting } from './components/admin/notifications/NotificationsRouting';
import { Entities } from './components/admin/notifications/Entities';
import { Components } from './components/admin/customActions/Components';
import { Actions } from './components/admin/customActions/Actions';
import { Connectors } from './components/admin/notifications/Connectors';
import { ApplicationConfig } from './components/admin/applicationSettings/ApplicationConfig';
import { JobComponents } from './components/admin/jobSettings/Components';
import { JobSettings } from './components/admin/jobSettings/JobSettings';
import { JobHistoryPage } from './components/admin/jobSettings/JobHistoryPage';
import { SearchConfig } from './components/admin/search/SearchConfig';
import { EndpointList } from './components/microsoft365/endpoints/EndpointList';
import { HomeDashboard } from './components/microsoft365/dashboards/Home';
import { PublicDashboard } from './components/microsoft365/dashboards/PublicAnnouncements';
import { ServiceHealth } from './components/microsoft365/health/ServiceHealth';
import { D365PPReleasesPage } from './components/microsoft365/d365pp/D365PPReleasesPage';

import { AzureLayout } from './components/azure/Layout';
import { AzureHomeDashboard } from './components/azure/dashboards/Home';
import { AzureServiceHealth } from './components/azure/health/ServiceHealth';
import { AzureUpdates } from './components/azure/updates/AzureUpdates';

import { ThemeProvider, ITheme } from '@fluentui/react';
import { FluentProvider, webLightTheme, teamsLightTheme, teamsDarkTheme, teamsHighContrastTheme, Theme } from '@fluentui/react-components';
import { M365ActualLightTheme } from '@m365-admin/customizations';
import { TeamsActualLightTheme, TeamsActualDarkTheme } from '@m365-admin/teams-theme';
import { getTeamsTheme } from './components/auth/detectTeams';
import { AccessGate } from './components/AccessGate';

interface IUserDashboards {
    [itemKey: string]: IRearrangeableGridDefinition;
}

export default class App extends Component<{}, IGlobalState>  {
    static displayName = App.name;

    setProfile = (profile: any) => {
        this.setState({
            userProfile: profile
        })
    }

    getProfileProperty = (name: string, defaultValue?: any): any => {
        return this.state.userProfile && this.state.userProfile[name] ?
            this.state.userProfile[name] : (defaultValue ? defaultValue : null);
    }

    setProfileProperty = (name: string, value: any) => {
        setPropertyAsync(name, value)
            .then((result: any) => {
                if (result.ok) {
                    var profile: any = this.state.userProfile ? this.state.userProfile : {};
                    profile[name] = value;
                    this.setState({
                        userProfile: profile
                    });
                }
            });
    }

    openNotificationsPanel = (open: boolean) => {
        this.setState({
            isNotificationsPanelOpen: open
        });
    }

    clearNotifications = () => {
        this.setState({
            notifications: []
        });

        sessionStorage.setItem('serviceHealthHubNotifications', JSON.stringify([]));
    }

    setToastNotification = (notification: INotification, skipToastNotification?: boolean): string => {
        var notifications: INotification[] = this.state.notifications;
        var existingIndex = notifications.findIndex((n) => n.id === notification.id);
        if (existingIndex !== -1) {
            // update notification
            notifications = notifications.filter((n) => n.id !== notification.id);
            notifications.push(notification);
        } else {
            // create unique id and push the notification
            notification.id = this._createGuid();
            notifications.push(notification);
        }

        notifications = notifications.sort((a, b) => a.time > b.time ? -1 : 1);

        sessionStorage.setItem('serviceHealthHubNotifications', JSON.stringify(notifications));

        if (skipToastNotification)
            this.setState({
                notifications: notifications
            })
        else {
            this.setState({
                toastNotification: notification,
                notifications: notifications
            });

            setTimeout(() => {
                this.setState({
                    toastNotification: undefined
                });
            }, 5000);
        }

        return notification.id;
    }

    clearToastNotification = () => {
        this.setState({
            toastNotification: undefined
        });
    }

    removeNotification = (notificationId: string): void => {
        var notifications: INotification[] = this.state.notifications;
        var existingIndex = notifications.findIndex((n) => n.id === notificationId);
        if (existingIndex !== -1) {
            // remove notification
            notifications = notifications.filter((n) => n.id !== notificationId);
        }

        notifications = notifications.sort((a, b) => a.time > b.time ? -1 : 1);
        sessionStorage.setItem('serviceHealthHubNotifications', JSON.stringify(notifications));

        this.setState({
            notifications: notifications
        });
    }

    getTheme = (): ITheme => {
        return this.state.theme;
    }

    getDashboard = (name: string): IRearrangeableGridDefinition | undefined => {
        var dashboards: IUserDashboards | undefined = this.getProfileProperty("dashboards", undefined);

        if (dashboards === undefined || dashboards === null)
            return undefined;
        else
            return dashboards[name];
    }

    setDashboard = (name: string, dashboard: IRearrangeableGridDefinition): void => {
        var dashboards: IUserDashboards | undefined = this.getProfileProperty("dashboards", undefined);

        if (dashboards === undefined || dashboards === null)
            dashboards = {
                [name]: dashboard
            }
        else
            dashboards[name] = dashboard;

        this.setProfileProperty("dashboards", dashboards);
    }

    constructor(props: {}) {
        super(props);

        var theme: any = M365ActualLightTheme;
        var fluentUI9Theme: Partial<Theme> = webLightTheme;

        switch (getTeamsTheme()) {
            case "default":
                theme = TeamsActualLightTheme;
                fluentUI9Theme = teamsLightTheme;
                break;
            case "dark":
                theme = TeamsActualDarkTheme;
                fluentUI9Theme = teamsDarkTheme;
                break;
            case "contrast":
                theme = TeamsActualLightTheme;
                fluentUI9Theme = teamsHighContrastTheme;
                break;
            default:
                theme = M365ActualLightTheme;
                fluentUI9Theme = webLightTheme;
        }

        this.state = {
            userProfile: {},
            notifications: [],
            toastNotification: undefined,
            isNotificationsPanelOpen: false,
            initialized: false,
            theme: theme,
            fluentUI9Theme: fluentUI9Theme,
            setProfile: this.setProfile,
            getProfileProperty: this.getProfileProperty,
            setProfileProperty: this.setProfileProperty,
            openNotificationsPanel: this.openNotificationsPanel,
            clearNotifications: this.clearNotifications,
            setToastNotification: this.setToastNotification,
            clearToastNotification: this.clearToastNotification,
            removeNotification: this.removeNotification,
            getTheme: this.getTheme,
            getDashboard: this.getDashboard,
            setDashboard: this.setDashboard
        }
    }

    render() {
        const { initialized } = this.state;
        return (
            <GlobalState.Provider value={this.state}>
                {initialized ? (
                    <FluentProvider theme={this.state.fluentUI9Theme}>
                        <ThemeProvider theme={this.state.theme} >
                            <Switch>
                                <Route exact path='/accessDenied' >
                                    <AccessDenied />
                                </Route>

                                <Route exact path='/auth-start' >
                                    <AuthStart />
                                </Route>
                                <Route exact path='/auth-end' >
                                    <AuthEnd />
                                </Route>

                                <Route exact path='/configure' >
                                    <Configure />
                                </Route>

                                <AuthenticationTemplate>
                                    <UserProfile>
                                        <ThemeProvider theme={this.state.theme} >
                                            <Route exact path='/admin'>
                                                <AccessGate requiredRoles={["Admin"]}>
                                                    <AdminHome />
                                                </AccessGate>
                                            </Route>
                                            <Route exact path='/admin/routing'>
                                                <AccessGate requiredRoles={["Admin"]}>
                                                    <AdminRouting />
                                                </AccessGate>
                                            </Route>
                                            <Route exact path='/admin/mapping'>
                                                <AccessGate requiredRoles={["Admin"]}>
                                                    <TaskMappingEntities />
                                                </AccessGate>
                                            </Route>
                                            <Route exact path='/admin/mapping/editor'>
                                                <AccessGate requiredRoles={["Admin"]}>
                                                    <AdminTaskMapping />
                                                </AccessGate>
                                            </Route>
                                            <Route exact path='/admin/notifications'>
                                                <AccessGate requiredRoles={["Admin"]}>
                                                    <NotificationsConfig />
                                                </AccessGate>
                                            </Route>
                                            <Route exact path='/admin/notifications/routing'>
                                                <AccessGate requiredRoles={["Admin"]}>
                                                    <NotificationsRouting />
                                                </AccessGate>
                                            </Route>
                                            <Route exact path='/admin/notifications/connectors'>
                                                <AccessGate requiredRoles={["Admin"]}>
                                                    <Connectors />
                                                </AccessGate>
                                            </Route>
                                            <Route exact path='/admin/notifications/entities'>
                                                <Entities />
                                            </Route>
                                            <Route exact path='/admin/customActions/components'>
                                                <AccessGate requiredRoles={["Admin"]}>
                                                    <Components />
                                                </AccessGate>
                                            </Route>
                                            <Route exact path='/admin/customActions/actions'>
                                                <AccessGate requiredRoles={["Admin"]}>
                                                    <Actions />
                                                </AccessGate>
                                            </Route>
                                            <Route exact path='/admin/settings'>
                                                <AccessGate requiredRoles={["Admin"]}>
                                                    <ApplicationConfig />
                                                </AccessGate>
                                            </Route>
                                            <Route exact path='/admin/jobs'>
                                                <AccessGate requiredRoles={["Admin"]}>
                                                    <JobComponents />
                                                </AccessGate>
                                            </Route>
                                            <Route exact path='/admin/jobs/settings'>
                                                <AccessGate requiredRoles={["Admin"]}>
                                                    <JobSettings />
                                                </AccessGate>
                                            </Route>
                                            <Route exact path='/admin/tags'>
                                                <AccessGate requiredRoles={["Admin"]}>
                                                    <TagDefinitionsList />
                                                </AccessGate>
                                            </Route>
                                            <Route exact path='/admin/jobs/history'>
                                                <AccessGate requiredRoles={["Admin"]}>
                                                    <JobHistoryPage />
                                                </AccessGate>
                                            </Route>
                                            <Route exact path='/admin/search'>
                                                <AccessGate requiredRoles={["Admin"]}>
                                                    <SearchConfig />
                                                </AccessGate>
                                            </Route>


                                            <Route exact path='/'>
                                                <Layout>
                                                    <AccessGate requiredRoles={["Admin", "ServiceHealthReader", "Communication.Write.All"]}>
                                                        <HomeDashboard />
                                                    </AccessGate>
                                                </Layout>
                                            </Route>
                                            <Route exact path='/messages'>
                                                <Layout>
                                                    <AccessGate requiredRoles={["Admin", "ServiceHealthReader", "Communication.Write.All"]}>
                                                        <MessagesList />
                                                    </AccessGate>
                                                </Layout>
                                            </Route>
                                            <Route exact path='/status'>
                                                <Layout>
                                                    <AccessGate requiredRoles={["Admin", "ServiceHealthReader", "Communication.Write.All"]}>
                                                        <ServiceDetailsList />
                                                    </AccessGate>
                                                </Layout>
                                            </Route>
                                            <Route exact path='/microsoft365/health'>
                                                <Layout>
                                                    <AccessGate requiredRoles={["Admin", "ServiceHealthReader", "Communication.Write.All"]}>
                                                        <ServiceHealth />
                                                    </AccessGate>
                                                </Layout>
                                            </Route>
                                            <Route exact path='/d365pp'>
                                                <Layout>
                                                    <AccessGate requiredRoles={["Admin", "ServiceHealthReader", "Communication.Write.All"]}>
                                                        <D365PPReleasesPage />
                                                    </AccessGate>
                                                </Layout>
                                            </Route>
                                            <Route exact path='/roadmap'>
                                                <Layout>
                                                    <AccessGate requiredRoles={["Admin", "ServiceHealthReader", "Communication.Write.All"]}>
                                                        <Roadmap />
                                                    </AccessGate>
                                                </Layout>
                                            </Route>
                                            <Route exact path='/reports/sharepoint'>
                                                <Layout>
                                                    <AccessGate requiredRoles={["Admin", "ServiceHealthReader", "Communication.Write.All"]}>
                                                        <SharePointOnlineReport />
                                                    </AccessGate>
                                                </Layout>
                                            </Route>
                                            <Route exact path='/reports/onedrive'>
                                                <Layout>
                                                    <AccessGate requiredRoles={["Admin", "ServiceHealthReader", "Communication.Write.All"]}>
                                                        <ODBOnlineReport />
                                                    </AccessGate>
                                                </Layout>
                                            </Route>
                                            <Route exact path='/reports/exchange'>
                                                <Layout>
                                                    <AccessGate requiredRoles={["Admin", "ServiceHealthReader", "Communication.Write.All"]}>
                                                        <ExchangeOnlineReport />
                                                    </AccessGate>
                                                </Layout>
                                            </Route>
                                            <Route exact path='/reports/skype'>
                                                <Layout>
                                                    <AccessGate requiredRoles={["Admin", "ServiceHealthReader", "Communication.Write.All"]}>
                                                        <SFBOnlineReport />
                                                    </AccessGate>
                                                </Layout>
                                            </Route>
                                            <Route exact path='/reports/teams'>
                                                <Layout>
                                                    <AccessGate requiredRoles={["Admin", "ServiceHealthReader", "Communication.Write.All"]}>
                                                        <MicrosoftTeamsOnlineReport />
                                                    </AccessGate>
                                                </Layout>
                                            </Route>
                                            <Route exact path='/reports/m365suite'>
                                                <Layout>
                                                    <AccessGate requiredRoles={["Admin", "ServiceHealthReader", "Communication.Write.All"]}>
                                                        <M365SuiteReport />
                                                    </AccessGate>
                                                </Layout>
                                            </Route>
                                            <Route exact path='/reports/dynamics'>
                                                <Layout>
                                                    <AccessGate requiredRoles={["Admin", "ServiceHealthReader", "Communication.Write.All"]}>
                                                        <Dynamics365Report />
                                                    </AccessGate>
                                                </Layout>
                                            </Route>
                                            <Route exact path='/reports/planner'>
                                                <Layout>
                                                    <AccessGate requiredRoles={["Admin", "ServiceHealthReader", "Communication.Write.All"]}>
                                                        <PlannerReport />
                                                    </AccessGate>
                                                </Layout>
                                            </Route>
                                            <Route exact path='/reports'>
                                                <Layout>
                                                    <AccessGate requiredRoles={["Admin", "ServiceHealthReader", "Communication.Write.All"]}>
                                                        <IncidentReport />
                                                    </AccessGate>
                                                </Layout>
                                            </Route>
                                            <Route exact path='/licenses'>
                                                <Layout>
                                                    <AccessGate requiredRoles={["Admin", "LicenseReader"]}>
                                                        <LicenseReport />
                                                    </AccessGate>
                                                </Layout>
                                            </Route>
                                            <Route exact path='/public'>
                                                <Layout hideMenu={true} >
                                                    <AccessGate requiredRoles={["Admin", "ServiceHealthReader", "Communication.Write.All", "Public", "LicenseReader"]}>
                                                        <PublicDashboard />
                                                    </AccessGate>
                                                </Layout>
                                            </Route>
                                            <Route exact path='/microsoft365/endpoints'>
                                                <Layout>
                                                    <AccessGate requiredRoles={["Admin", "ServiceHealthReader", "Communication.Write.All"]}>
                                                        <EndpointList />
                                                    </AccessGate>
                                                </Layout>
                                            </Route>
                                            <Route exact path='/settings'>
                                                <Layout>
                                                    <AccessGate requiredRoles={["Admin", "ServiceHealthReader", "Communication.Write.All", "Public", "LicenseReader"]}>
                                                        <Settings />
                                                    </AccessGate>
                                                </Layout>
                                            </Route>

                                            <Route exact path='/azure'>
                                                <AzureLayout>
                                                    <AccessGate requiredRoles={["Admin", "ServiceHealthReader", "Communication.Write.All"]}>
                                                        <AzureHomeDashboard />
                                                    </AccessGate>
                                                </AzureLayout>
                                            </Route>
                                            <Route exact path='/azure/health'>
                                                <AzureLayout>
                                                    <AccessGate requiredRoles={["Admin", "ServiceHealthReader", "Communication.Write.All"]}>
                                                        <AzureServiceHealth />
                                                    </AccessGate>
                                                </AzureLayout>
                                            </Route>
                                            <Route exact path='/azure/updates'>
                                                <AzureLayout>
                                                    <AccessGate requiredRoles={["Admin", "ServiceHealthReader", "Communication.Write.All"]}>
                                                        <AzureUpdates />
                                                    </AccessGate>
                                                </AzureLayout>
                                            </Route>
                                        </ThemeProvider>
                                    </UserProfile>
                                </AuthenticationTemplate>
                            </Switch>
                        </ThemeProvider>
                    </FluentProvider>
                ) : (<div />)}
            </GlobalState.Provider>
        );
    }

    componentDidMount() {
        this.setTheme();
        this.initialize();
    }

    initialize() {
        const notifications: INotification[] = this._getNotifications();
        this.setState({
            notifications: notifications,
            initialized: true
        });
    }

    private setTheme() {
        // eslint-disable-next-line dot-notation
        const microsoftTeamsLib = microsoftTeams || window["microsoftTeams"];

        if (!microsoftTeamsLib) {
            this.setState({
                theme: M365ActualLightTheme
            });
        }

        microsoftTeams.getContext((context: microsoftTeams.Context) => {
            this.setThemeInt(context.theme ? context.theme : "notInTeams");
        });

        microsoftTeams.registerOnThemeChangeHandler((theme: string) => {
            this.setThemeInt(theme);
        });
    };

    private setThemeInt(theme: string) {
        var fluentUITheme: ITheme = M365ActualLightTheme;

        switch (theme) {
            case "default":
                fluentUITheme = TeamsActualLightTheme;
                break;
            case "dark":
                fluentUITheme = TeamsActualDarkTheme;
                break;
            case "contrast":
                fluentUITheme = TeamsActualLightTheme;
                break;
            default:
                fluentUITheme = M365ActualLightTheme;
        }

        this.setState({
            theme: fluentUITheme
        });
    }

    private _getNotifications(): INotification[] {
        var notifications: INotification[] = [];
        var sessionStorageNotificationsStr: string | null | undefined = sessionStorage.getItem('serviceHealthHubNotifications');

        if (sessionStorageNotificationsStr && sessionStorageNotificationsStr !== null) {
            var sessionStorageNotifications: any = JSON.parse(sessionStorageNotificationsStr);
            for (const notification of sessionStorageNotifications)
                notifications.push({
                    id: notification.id,
                    state: notification.state,
                    title: notification.title,
                    message: notification.message,
                    time: new Date(notification.time),
                    active: notification.active
                });
        }

        return notifications;
    }

    private _createGuid(): string {
        return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
            var r = Math.random() * 16 | 0, v = c === 'x' ? r : (r & 0x3 | 0x8);
            return v.toString(16);
        });
    }
}