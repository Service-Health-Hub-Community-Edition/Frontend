import * as React from 'react';
import { Nav, INavLinkGroup } from '@m365-admin/nav';

interface INavAdminState {
    windowWidth: number;
    sidebarCollapsed: boolean;
    isNotificationsGroupExpanded: boolean;
}

export class NavAdmin extends React.Component<{}, INavAdminState> {
    constructor(props: { }) {
        super(props);

        this.state = {
            windowWidth: this._getWidth(),
            sidebarCollapsed: false,
            isNotificationsGroupExpanded: false
        };

        window.addEventListener('resize', this._handleResize);
    }

    public render() {
        const relativeUrlPath: string = this._getUrlRelativePath().toLowerCase();
        const { windowWidth } = this.state;

        const adminNav: INavLinkGroup[] = [
            {
                key: 'shhAdmin',
                name: 'Admin Center',
                links: [
                    {
                        name: 'Home',
                        icon: 'Home',
                        key: 'shhAdminHome',
                        href: '/admin',
                        isSelected: relativeUrlPath === '/admin',
                        style: { textDecoration: 'none' }
                    },
                    {
                        name: 'Notifications',
                        icon: 'News',
                        key: 'shhAdminNotificationsGroup',
                        style: { textDecoration: 'none' },
                        links: [
                            {
                                name: 'Get started',
                                key: 'shhAdminNotificationGetStarted',
                                href: '/admin/notifications',
                                isSelected: relativeUrlPath === '/admin/notifications',
                                style: { textDecoration: 'none' }
                            },
                            {
                                name: 'Connectors',
                                key: 'shhAdminNotificationConnectors',
                                href: '/admin/notifications/connectors',
                                isSelected: relativeUrlPath === '/admin/notifications/connectors',
                                style: { textDecoration: 'none' }
                            },
                            {
                                name: 'Routing',
                                key: 'shhAdminNotificationEntities',
                                href: '/admin/notifications/entities',
                                isSelected: relativeUrlPath === '/admin/notifications/entities',
                                style: { textDecoration: 'none' }
                            },
                        ]
                    },
                    {
                        name: 'Jobs',
                        icon: 'AutoRacing',
                        key: 'shhAdminJobs',
                        href: '/admin/jobs',
                        isSelected: relativeUrlPath === '/admin/jobs',
                        style: { textDecoration: 'none' }
                    },
                    {
                        name: 'Mapping',
                        icon: 'DataflowsLink',
                        key: 'shhAdminMapping',
                        href: '/admin/mapping',
                        isSelected: relativeUrlPath === '/admin/mapping',
                        style: { textDecoration: 'none' }
                    },
                    {
                        name: 'Integration',
                        icon: 'AzureServiceEndpoint',
                        key: 'shhAdminCustomActions',
                        href: '/admin/customActions/components',
                        isSelected: relativeUrlPath.startsWith('/admin/customactions'),
                        style: { textDecoration: 'none' }
                    },
                    {
                        name: 'Tags',
                        icon: 'Tag',
                        key: 'shhAdminTags',
                        href: '/admin/tags',
                        isSelected: relativeUrlPath.startsWith('/admin/tags'),
                        style: { textDecoration: 'none' }
                    },
                    {
                        name: 'Search',
                        icon: 'Search',
                        key: 'shhAdminSearch',
                        href: '/admin/search',
                        isSelected: relativeUrlPath.startsWith('/admin/search'),
                        style: { textDecoration: 'none' }
                    },
                    {
                        name: 'Settings',
                        icon: 'Settings',
                        key: 'shhAdminSettings',
                        href: '/admin/settings',
                        isSelected: relativeUrlPath.startsWith('/admin/settings'),
                        style: { textDecoration: 'none' }
                    }
                ]
            }
        ];

        return (
                <Nav 
                    groups={adminNav}
                    enableCustomization={false}
                showMore={true}
                styles={{ root: { textDecoration: 'inherit', height: "100%" }, navContainer: { paddingRight: "8px" } }}
                    defaultIsNavCollapsed={windowWidth > 500 ? false : true} 
                />
        );
    }
    
    private _getUrlRelativePath = (): string => {
        var url = document.location.toString();
        var arrUrl = url.split("//");

        var start = arrUrl[1].indexOf("/");
        var relUrl = arrUrl[1].substring(start);// "stop" is omitted, intercept all characters from start to end

        if (relUrl.indexOf("?") != -1) {
            relUrl = relUrl.split("?")[0];
        }
        return relUrl.toLowerCase();
    }

    private _handleResize = () => {
        this.setState({
            windowWidth: this._getWidth()
        });
    }

    private _getWidth() {
        return window.innerWidth && document.documentElement.clientWidth ?
            Math.min(window.innerWidth, document.documentElement.clientWidth) :
            window.innerWidth ||
            document.documentElement.clientWidth ||
            document.getElementsByTagName('body')[0].clientWidth;
    }
}