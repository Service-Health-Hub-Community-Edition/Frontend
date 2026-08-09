import * as React from 'react';
import { Nav, INavLinkGroup, INavLink } from '@m365-admin/nav';
import { acquireAccessToken } from '../../auth/AccessTokenHelper';

interface INavAzureState {
    windowWidth: number;
    navItems: INavLink[];
    sidebarCollapsed: boolean;
    isNotificationsGroupExpanded: boolean;
}

interface IServiceHubNavLink {
    key: string,
    icon?: string,
    name: string,
    href?: string,
    style?: React.CSSProperties,
    isSelected?: boolean;
    roles: string[],
    links?: IServiceHubNavLink[]
}

export class NavAzure extends React.Component<{}, INavAzureState> {
    constructor(props: { }) {
        super(props);

        this.state = {
            windowWidth: this._getWidth(),
            navItems: [],
            sidebarCollapsed: false,
            isNotificationsGroupExpanded: false
        };
 
        window.addEventListener('resize', this._handleResize);
    }

    public render() {
        
        const { windowWidth, navItems } = this.state;

        const adminNav: INavLinkGroup[] = [
            {
                key: 'shhAzure',
                name: 'Microsoft Azure',
                links: navItems
            }
        ];

        return (
                <Nav 
                    groups={adminNav}
                    enableCustomization={false}
                    showMore={true}
                styles={{ root: { textDecoration: 'inherit', height: "100%" }, navContainer: { paddingRight: "8px" } }}
                    defaultIsNavCollapsed={true} 
                />
        );
    }

    componentDidMount() {
        this._handleResize();

        var navLinkGroups: INavLink[] = [];
        const relativeUrlPath: string = this._getUrlRelativePath().toLowerCase();

        const seriviceHubNavLinkGroups: IServiceHubNavLink[] = [
                {
                    name: 'Home',
                    icon: 'Home',
                    key: 'shhAzureHome',
                    href: '/azure',
                    isSelected: relativeUrlPath === '/azure',
                    style: { textDecoration: 'none' },
                    roles: ['Admin', 'ServiceHealthReader', 'Communication.Write.All']
                },
                {
                    name: 'Service health',
                    icon: 'Health',
                    key: 'shhAzureHealth',
                    href: '/azure/health',
                    isSelected: relativeUrlPath === '/azure/health',
                    style: { textDecoration: 'none' },
                    roles: ['Admin', 'ServiceHealthReader', 'Communication.Write.All']
                },
                {
                    name: 'Azure updates',
                    icon: 'News',
                    key: 'shhAzureUpdates',
                    href: '/azure/updates',
                    isSelected: relativeUrlPath === '/azure/updates',
                    style: { textDecoration: 'none' },
                    roles: ['Admin', 'ServiceHealthReader', 'Communication.Write.All']
                },
            ];

        acquireAccessToken()
            .then((result) => {
                var tokenClaims: any = result.account?.idTokenClaims;
                const roles: any = tokenClaims?.roles;

                if (roles !== undefined) {
                    for (const navGroup of seriviceHubNavLinkGroups) {
                        if (roles.some((r: string) => navGroup.roles.includes(r))) {
                            if (navGroup.links === undefined) {
                                navLinkGroups.push({
                                    key: navGroup.key,
                                    name: navGroup.name,
                                    href: navGroup.href,
                                    isSelected: navGroup.isSelected,
                                    style: navGroup.style,
                                    icon: navGroup.icon
                                });
                            } else {
                                var navLinks: IServiceHubNavLink[] = [];

                                for (const link of navGroup.links) {
                                    if (roles.some((r: string) => link.roles.includes(r))) {
                                        navLinks.push(link);
                                    }
                                }

                                if (navLinks.length > 0) {
                                    var ngLinks: INavLink[] = [];
                                    for (const nLink of navLinks) {
                                        ngLinks.push({
                                            key: nLink.key,
                                            name: nLink.name,
                                            href: nLink.href,
                                            isSelected: nLink.isSelected,
                                            style: nLink.style,
                                            icon: nLink.icon
                                        });
                                    }

                                    navLinkGroups.push({
                                        key: navGroup.key,
                                        name: navGroup.name,
                                        href: navGroup.href,
                                        isSelected: navGroup.isSelected,
                                        style: navGroup.style,
                                        icon: navGroup.icon,
                                        links: ngLinks
                                    });
                                }
                            }
                        }
                    }
                }

                this.setState({
                    navItems: navLinkGroups
                });
            });
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