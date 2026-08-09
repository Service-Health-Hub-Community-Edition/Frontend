import * as React from 'react';
import { Nav, INavLinkGroup, INavLink } from '@m365-admin/nav';
import { checkInTeams } from '../auth/detectTeams';
import { acquireAccessToken } from '../../auth/AccessTokenHelper';
import { ThemeProvider, ITheme } from '@fluentui/react';
import { GlobalState } from '../GlobalState';

interface INavM365State {
    windowWidth: number;
    navItems: INavLink[];
    sidebarCollapsed: boolean;
    isNotificationsGroupExpanded: boolean;
    theme: any;
    initialized: boolean;
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

interface IServiceHubNavLinkGroup {
    name: string,
    expandAriaLabel: string,
    collapseAriaLabel: string,
    links: IServiceHubNavLink[]
}

export class NavM365 extends React.Component<{}, INavM365State> {
    static contextType = GlobalState;
    constructor(props: { }) {
        super(props);

        this.state = {
            windowWidth: this._getWidth(),
            navItems: [],
            sidebarCollapsed: false,
            isNotificationsGroupExpanded: false,
            theme: undefined,
            initialized: false
        };
 
        window.addEventListener('resize', this._handleResize);
    }

    public render() {
        
        const { windowWidth, navItems, initialized } = this.state;

        const adminNav: INavLinkGroup[] = [
            {
                key: 'shhM365',
                name: 'Microsoft 365',
                links: navItems
            }
        ];

        if (!initialized) {
            return (<></>);
        }

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

        let globalState: any = this.context;
        var theme: ITheme = globalState.getTheme();

        this.setState({
            theme: theme
        });

        var navLinkGroups: INavLink[] = [];
        const relativeUrlPath: string = this._getUrlRelativePath().toLowerCase();

        const seriviceHubNavLinkGroups: IServiceHubNavLink[] = [
                {
                    name: 'Home',
                    icon: 'Home',
                    key: 'shhHome',
                    href: '/',
                    isSelected: relativeUrlPath === '/',
                    style: { textDecoration: 'none' },
                roles: ['Admin', 'Communication.Write.All', 'ServiceHealthReader']
                },
                {
                    name: 'Service health',
                    icon: 'Health',
                    key: 'shhHealth',
                    href: '/microsoft365/health',
                    isSelected: relativeUrlPath === '/microsoft365/health',
                    style: { textDecoration: 'none' },
                    roles: ['Admin', 'Communication.Write.All', 'ServiceHealthReader']
                },
                {
                    name: 'Reports',
                    icon: 'PowerBILogo16',
                    key: 'shhReportsGroup',
                    style: { textDecoration: 'none' },
                    roles: ['Admin', 'Communication.Write.All', 'ServiceHealthReader', 'LicenseReader'],
                    links: [
                        {
                            name: 'Incident reports',
                            key: 'shhIncidentReports',
                            href: '/reports',
                            isSelected: relativeUrlPath.startsWith('/reports'),
                            style: { textDecoration: 'none' },
                            roles: ['Admin', 'ServiceHealthReader', 'Communication.Write.All']
                        },
                        {
                            name: 'Licenses',
                            key: 'shhLicenses',
                            href: '/licenses',
                            isSelected: relativeUrlPath === '/licenses',
                            style: { textDecoration: 'none' },
                            roles: ['Admin', 'LicenseReader']
                        }
                    ]
                },
                {
                    name: 'News',
                    icon: 'News',
                    key: 'shhNewsGroup',
                    style: { textDecoration: 'none' },
                    roles: ['Admin', 'Communication.Write.All', 'ServiceHealthReader'],
                    links: [
                        {
                            name: 'Message Center',
                            key: 'shhMessageCenter',
                            href: '/messages',
                            isSelected: relativeUrlPath === '/messages',
                            style: { textDecoration: 'none' },
                            roles: ['Admin', 'Communication.Write.All', 'ServiceHealthReader']
                        },
                        {
                            name: 'Roadmap',
                            key: 'shhRoadmap',
                            href: '/roadmap',
                            isSelected: relativeUrlPath === '/roadmap',
                            style: { textDecoration: 'none' },
                            roles: ['Admin', 'Communication.Write.All', 'ServiceHealthReader']
                        },
                        {
                            name: 'D365 and Power Platform',
                            key: 'shhD365PP',
                            href: '/d365pp',
                            isSelected: relativeUrlPath === '/d365pp',
                            style: { textDecoration: 'none' },
                            roles: ['Admin', 'Communication.Write.All', 'ServiceHealthReader']
                        }
                    ]
                },
                {
                    name: 'Office 365 Endpoints',
                    icon: 'WifiEthernet',
                    key: 'shhO365Endpoints',
                    href: '/microsoft365/endpoints',
                    isSelected: relativeUrlPath === '/microsoft365/endpoints',
                    style: { textDecoration: 'none' },
                    roles: ['Admin', 'Communication.Write.All', 'ServiceHealthReader']
                }
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
                    navItems: navLinkGroups,
                    initialized: true
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