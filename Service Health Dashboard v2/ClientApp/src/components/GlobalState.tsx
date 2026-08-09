import * as React from 'react';
import { NotificationState, INotification } from './Notifications';
import { ITheme } from '@fluentui/react';
import { Theme } from '@fluentui/react-components';
import { IM365Theme } from '@m365-admin/customizations';
import type { IRearrangeableGridDefinition } from '@m365-admin/rearrangeable-grid';

export interface IGlobalState {
    userProfile: any;
    notifications: INotification[];
    toastNotification: INotification | undefined;
    isNotificationsPanelOpen: boolean;
    initialized: boolean;
    theme: any;
    fluentUI9Theme: Partial<Theme>;
    setProfile(profile: any): void;
    getProfileProperty(name: string, defaultValue?: any): any;
    setProfileProperty(name: string, value: any): void;
    openNotificationsPanel(open: boolean): void;
    clearNotifications(): void;
    setToastNotification(notification: INotification, skipToastNotification?: boolean): string;
    clearToastNotification(): void;
    removeNotification(notificationId: string): void;
    getTheme(): ITheme;
    getDashboard(name: string): IRearrangeableGridDefinition | undefined;
    setDashboard(name: string, dashboard: IRearrangeableGridDefinition): void;
}

export const GlobalState = React.createContext({
    userProfile: {}
});