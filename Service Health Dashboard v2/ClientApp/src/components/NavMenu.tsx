import * as React from 'react';
import { Container, Navbar, NavbarBrand } from 'reactstrap';
import { GlobalState } from './GlobalState'; 
import { Link } from 'react-router-dom';
import { Text, IDropdownOption, IDropdownStyles, Button, ButtonType, Image, IconType, IButtonProps } from '@fluentui/react';
import { Panel, PanelType, Nav, INavStyles, INavLinkGroup, INavLink, IconButton, Overlay } from '@fluentui/react';
import { CounterBadge } from "@fluentui/react-components";
import { Notifications } from './Notifications';
import { IApplicationBrandingSettings } from './admin/applicationSettings/ApplicationConfig';
import './NavMenu.css';
import { checkInTeams } from './auth/detectTeams';
import { acquireAccessToken } from '../auth/AccessTokenHelper';

export interface INavMenuState {
    isMenuPanelOpen: boolean;
    collapsed: boolean;
    showLogo: boolean;
    inTeams: boolean;
    languageOptions: IDropdownOption[];
    navLinkGroups: INavLinkGroup[];
    branding: IApplicationBrandingSettings;
    appTitle: string;
}

interface IServiceHubNavLink {
    key: string,
    name: string,
    url: string,
    icon?: string,
    roles: string[]
}

interface IServiceHubNavLinkGroup {
    name: string,
    expandAriaLabel: string,
    collapseAriaLabel: string,
    links: IServiceHubNavLink[]
}

export class NavMenu extends React.Component<{}, INavMenuState> {
    static displayName = NavMenu.name;
    static contextType = GlobalState;

    constructor(props: {}) {
        super(props);

        this.state = {
            isMenuPanelOpen: false,
            collapsed: true,
            showLogo: window.innerWidth < 500 ? false : true,
            inTeams: checkInTeams(),
            languageOptions: [],
            navLinkGroups: [],
            branding: {
                enabled: false,
                backgroundColor: '',
                logo: ''
            },
            appTitle: 'Service Health Hub'
        };

        window.addEventListener('resize', this._handleResize);
    }

    render() {
        const { isMenuPanelOpen, languageOptions, navLinkGroups, showLogo, branding, appTitle } = this.state;
        let globalState: any = this.context;

        const navStyles: Partial<INavStyles> = { root: { maxWidth: 250 } };

        const dropdownStyles: Partial<IDropdownStyles> = {
            dropdown: { width: 250 },
        };

        return (
            this.state.inTeams ? "" : 
                <header>
                    <Navbar className="shh-header-nav-root-no-side-padding" dark >
                        <Container style={{ padding: '0px', overflow: 'hidden' }}>
                        <div className="container-fluid" style={{ padding: '0px'}}>
                            <div className="row-fluid d-flex" style={{ width: '100%' }}>
                                <div className="col-auto" style={{ textAlign: 'left', display: 'inline-block', justifyContent: 'center' }}>
                                    <svg className="waffleMenu" xmlns="http://www.w3.org/2000/svg" onClick={() => this._onOpenMenuPanel()}>
                                        <path d="M1.6 0C1.82222 0 2.03056 0.0416667 2.225 0.125C2.41944 0.208333 2.58889 0.322222 2.73333 0.466667C2.87778 0.611111 2.99167 0.780556 3.075 0.975C3.15833 1.16944 3.2 1.37778 3.2 1.6C3.2 1.82222 3.15833 2.03056 3.075 2.225C2.99167 2.41944 2.87778 2.58889 2.73333 2.73333C2.58889 2.87778 2.41944 2.99167 2.225 3.075C2.03056 3.15833 1.82222 3.2 1.6 3.2C1.37778 3.2 1.16944 3.15833 0.975 3.075C0.780556 2.99167 0.611111 2.87778 0.466667 2.73333C0.322222 2.58889 0.208333 2.41944 0.125 2.225C0.0416667 2.03056 0 1.82222 0 1.6C0 1.37778 0.0416667 1.16944 0.125 0.975C0.208333 0.780556 0.322222 0.611111 0.466667 0.466667C0.611111 0.322222 0.780556 0.208333 0.975 0.125C1.16944 0.0416667 1.37778 0 1.6 0ZM8 0C8.22222 0 8.43056 0.0416667 8.625 0.125C8.81944 0.208333 8.98889 0.322222 9.13333 0.466667C9.27778 0.611111 9.39167 0.780556 9.475 0.975C9.55833 1.16944 9.6 1.37778 9.6 1.6C9.6 1.82222 9.55833 2.03056 9.475 2.225C9.39167 2.41944 9.27778 2.58889 9.13333 2.73333C8.98889 2.87778 8.81944 2.99167 8.625 3.075C8.43056 3.15833 8.22222 3.2 8 3.2C7.77778 3.2 7.56944 3.15833 7.375 3.075C7.18056 2.99167 7.01111 2.87778 6.86667 2.73333C6.72222 2.58889 6.60833 2.41944 6.525 2.225C6.44167 2.03056 6.4 1.82222 6.4 1.6C6.4 1.37778 6.44167 1.16944 6.525 0.975C6.60833 0.780556 6.72222 0.611111 6.86667 0.466667C7.01111 0.322222 7.18056 0.208333 7.375 0.125C7.56944 0.0416667 7.77778 0 8 0ZM14.4 3.2C14.1778 3.2 13.9694 3.15833 13.775 3.075C13.5806 2.99167 13.4111 2.87778 13.2667 2.73333C13.1222 2.58889 13.0083 2.41944 12.925 2.225C12.8417 2.03056 12.8 1.82222 12.8 1.6C12.8 1.37778 12.8417 1.16944 12.925 0.975C13.0083 0.780556 13.1222 0.611111 13.2667 0.466667C13.4111 0.322222 13.5806 0.208333 13.775 0.125C13.9694 0.0416667 14.1778 0 14.4 0C14.6222 0 14.8306 0.0416667 15.025 0.125C15.2194 0.208333 15.3889 0.322222 15.5333 0.466667C15.6778 0.611111 15.7917 0.780556 15.875 0.975C15.9583 1.16944 16 1.37778 16 1.6C16 1.82222 15.9583 2.03056 15.875 2.225C15.7917 2.41944 15.6778 2.58889 15.5333 2.73333C15.3889 2.87778 15.2194 2.99167 15.025 3.075C14.8306 3.15833 14.6222 3.2 14.4 3.2ZM1.6 6.4C1.82222 6.4 2.03056 6.44167 2.225 6.525C2.41944 6.60833 2.58889 6.72222 2.73333 6.86667C2.87778 7.01111 2.99167 7.18056 3.075 7.375C3.15833 7.56944 3.2 7.77778 3.2 8C3.2 8.22222 3.15833 8.43056 3.075 8.625C2.99167 8.81944 2.87778 8.98889 2.73333 9.13333C2.58889 9.27778 2.41944 9.39167 2.225 9.475C2.03056 9.55833 1.82222 9.6 1.6 9.6C1.37778 9.6 1.16944 9.55833 0.975 9.475C0.780556 9.39167 0.611111 9.27778 0.466667 9.13333C0.322222 8.98889 0.208333 8.81944 0.125 8.625C0.0416667 8.43056 0 8.22222 0 8C0 7.77778 0.0416667 7.56944 0.125 7.375C0.208333 7.18056 0.322222 7.01111 0.466667 6.86667C0.611111 6.72222 0.780556 6.60833 0.975 6.525C1.16944 6.44167 1.37778 6.4 1.6 6.4ZM8 6.4C8.22222 6.4 8.43056 6.44167 8.625 6.525C8.81944 6.60833 8.98889 6.72222 9.13333 6.86667C9.27778 7.01111 9.39167 7.18056 9.475 7.375C9.55833 7.56944 9.6 7.77778 9.6 8C9.6 8.22222 9.55833 8.43056 9.475 8.625C9.39167 8.81944 9.27778 8.98889 9.13333 9.13333C8.98889 9.27778 8.81944 9.39167 8.625 9.475C8.43056 9.55833 8.22222 9.6 8 9.6C7.77778 9.6 7.56944 9.55833 7.375 9.475C7.18056 9.39167 7.01111 9.27778 6.86667 9.13333C6.72222 8.98889 6.60833 8.81944 6.525 8.625C6.44167 8.43056 6.4 8.22222 6.4 8C6.4 7.77778 6.44167 7.56944 6.525 7.375C6.60833 7.18056 6.72222 7.01111 6.86667 6.86667C7.01111 6.72222 7.18056 6.60833 7.375 6.525C7.56944 6.44167 7.77778 6.4 8 6.4ZM14.4 6.4C14.6222 6.4 14.8306 6.44167 15.025 6.525C15.2194 6.60833 15.3889 6.72222 15.5333 6.86667C15.6778 7.01111 15.7917 7.18056 15.875 7.375C15.9583 7.56944 16 7.77778 16 8C16 8.22222 15.9583 8.43056 15.875 8.625C15.7917 8.81944 15.6778 8.98889 15.5333 9.13333C15.3889 9.27778 15.2194 9.39167 15.025 9.475C14.8306 9.55833 14.6222 9.6 14.4 9.6C14.1778 9.6 13.9694 9.55833 13.775 9.475C13.5806 9.39167 13.4111 9.27778 13.2667 9.13333C13.1222 8.98889 13.0083 8.81944 12.925 8.625C12.8417 8.43056 12.8 8.22222 12.8 8C12.8 7.77778 12.8417 7.56944 12.925 7.375C13.0083 7.18056 13.1222 7.01111 13.2667 6.86667C13.4111 6.72222 13.5806 6.60833 13.775 6.525C13.9694 6.44167 14.1778 6.4 14.4 6.4ZM1.6 12.8C1.82222 12.8 2.03056 12.8417 2.225 12.925C2.41944 13.0083 2.58889 13.1222 2.73333 13.2667C2.87778 13.4111 2.99167 13.5806 3.075 13.775C3.15833 13.9694 3.2 14.1778 3.2 14.4C3.2 14.6222 3.15833 14.8306 3.075 15.025C2.99167 15.2194 2.87778 15.3889 2.73333 15.5333C2.58889 15.6778 2.41944 15.7917 2.225 15.875C2.03056 15.9583 1.82222 16 1.6 16C1.37778 16 1.16944 15.9583 0.975 15.875C0.780556 15.7917 0.611111 15.6778 0.466667 15.5333C0.322222 15.3889 0.208333 15.2194 0.125 15.025C0.0416667 14.8306 0 14.6222 0 14.4C0 14.1778 0.0416667 13.9694 0.125 13.775C0.208333 13.5806 0.322222 13.4111 0.466667 13.2667C0.611111 13.1222 0.780556 13.0083 0.975 12.925C1.16944 12.8417 1.37778 12.8 1.6 12.8ZM8 12.8C8.22222 12.8 8.43056 12.8417 8.625 12.925C8.81944 13.0083 8.98889 13.1222 9.13333 13.2667C9.27778 13.4111 9.39167 13.5806 9.475 13.775C9.55833 13.9694 9.6 14.1778 9.6 14.4C9.6 14.6222 9.55833 14.8306 9.475 15.025C9.39167 15.2194 9.27778 15.3889 9.13333 15.5333C8.98889 15.6778 8.81944 15.7917 8.625 15.875C8.43056 15.9583 8.22222 16 8 16C7.77778 16 7.56944 15.9583 7.375 15.875C7.18056 15.7917 7.01111 15.6778 6.86667 15.5333C6.72222 15.3889 6.60833 15.2194 6.525 15.025C6.44167 14.8306 6.4 14.6222 6.4 14.4C6.4 14.1778 6.44167 13.9694 6.525 13.775C6.60833 13.5806 6.72222 13.4111 6.86667 13.2667C7.01111 13.1222 7.18056 13.0083 7.375 12.925C7.56944 12.8417 7.77778 12.8 8 12.8ZM14.4 12.8C14.6222 12.8 14.8306 12.8417 15.025 12.925C15.2194 13.0083 15.3889 13.1222 15.5333 13.2667C15.6778 13.4111 15.7917 13.5806 15.875 13.775C15.9583 13.9694 16 14.1778 16 14.4C16 14.6222 15.9583 14.8306 15.875 15.025C15.7917 15.2194 15.6778 15.3889 15.5333 15.5333C15.3889 15.6778 15.2194 15.7917 15.025 15.875C14.8306 15.9583 14.6222 16 14.4 16C14.1778 16 13.9694 15.9583 13.775 15.875C13.5806 15.7917 13.4111 15.6778 13.2667 15.5333C13.1222 15.3889 13.0083 15.2194 12.925 15.025C12.8417 14.8306 12.8 14.6222 12.8 14.4C12.8 14.1778 12.8417 13.9694 12.925 13.775C13.0083 13.5806 13.1222 13.4111 13.2667 13.2667C13.4111 13.1222 13.5806 13.0083 13.775 12.925C13.9694 12.8417 14.1778 12.8 14.4 12.8Z"></path>
                                    </svg>
                                </div>
                                    {showLogo ? (
                                        <div className="col-auto" style={{ textAlign: 'left', display: 'inline-block', justifyContent: 'center' }}>
                                            <img
                                                id="O365_MainLink_TenantLogoImg"
                                                height='24px'
                                                alt='Home'
                                                style={{ marginTop: '-6px', cursor: 'pointer' }}
                                                onClick={() => window.location.href = '/'}
                                                src={
                                                    branding.enabled && branding.logo !== undefined && branding.logo.trim() !== '' ?
                                                        branding.logo.trim() :
                                                        "data:image/svg+xml;base64,PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0idXRmLTgiPz4KPCEtLSBHZW5lcmF0b3I6IEFkb2JlIElsbHVzdHJhdG9yIDI0LjAuMywgU1ZHIEV4cG9ydCBQbHVnLUluIC4gU1ZHIFZlcnNpb246IDYuMDAgQnVpbGQgMCkgIC0tPgo8c3ZnIHZlcnNpb249IjEuMSIgaWQ9Ik1TLXN5bWJvbCIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIiB4bWxuczp4bGluaz0iaHR0cDovL3d3dy53My5vcmcvMTk5OS94bGluayIgeD0iMHB4IiB5PSIwcHgiCgkgdmlld0JveD0iMCAwIDMzNy42IDcyIiBzdHlsZT0iZW5hYmxlLWJhY2tncm91bmQ6bmV3IDAgMCAzMzcuNiA3MjsiIHhtbDpzcGFjZT0icHJlc2VydmUiPgo8c3R5bGUgdHlwZT0idGV4dC9jc3MiPgoJLnN0MHtmaWxsOiNGRkZGRkY7fQoJLnN0MXtmaWxsOiNGMjUwMjI7fQoJLnN0MntmaWxsOiM3RkJBMDA7fQoJLnN0M3tmaWxsOiMwMEE0RUY7fQoJLnN0NHtmaWxsOiNGRkI5MDA7fQo8L3N0eWxlPgo8Zz4KCTxwYXRoIGNsYXNzPSJzdDAiIGQ9Ik0xNDAuMywxNC40djQzLjJoLTcuNVYyMy43aC0wLjFsLTEzLjQsMzMuOWgtNWwtMTMuNy0zMy45aC0wLjF2MzMuOWgtNi45VjE0LjRoMTAuOGwxMi40LDMyaDAuMmwxMy4xLTMyCgkJSDE0MC4zeiBNMTQ2LjYsMTcuN2MwLTEuMiwwLjQtMi4yLDEuMy0zYzAuOS0wLjgsMS45LTEuMiwzLjEtMS4yYzEuMywwLDIuNCwwLjQsMy4yLDEuM2MwLjgsMC44LDEuMywxLjgsMS4zLDNzLTAuNCwyLjItMS4zLDMKCQljLTAuOSwwLjgtMS45LDEuMi0zLjIsMS4ycy0yLjMtMC40LTMuMS0xLjJDMTQ3LDE5LjgsMTQ2LjYsMTguOCwxNDYuNiwxNy43eiBNMTU0LjcsMjYuNnYzMWgtNy4zdi0zMQoJCUMxNDcuNCwyNi42LDE1NC43LDI2LjYsMTU0LjcsMjYuNnogTTE3Ni44LDUyLjNjMS4xLDAsMi4zLTAuMywzLjYtMC44czIuNS0xLjIsMy42LTJ2Ni44Yy0xLjIsMC43LTIuNSwxLjItNCwxLjVzLTMuMSwwLjUtNC45LDAuNQoJCWMtNC42LDAtOC4zLTEuNC0xMS4xLTQuM2MtMi45LTIuOS00LjMtNi42LTQuMy0xMWMwLTUsMS41LTkuMSw0LjQtMTIuM2MyLjktMy4yLDctNC44LDEyLjQtNC44YzEuNCwwLDIuNywwLjIsNC4xLDAuNQoJCWMxLjQsMC40LDIuNSwwLjgsMy4zLDEuMnY3Yy0xLjEtMC44LTIuMy0xLjUtMy40LTEuOWMtMS4yLTAuNS0yLjQtMC43LTMuNi0wLjdjLTIuOSwwLTUuMiwwLjktNywyLjhzLTIuNyw0LjQtMi43LDcuNgoJCWMwLDMuMSwwLjgsNS42LDIuNSw3LjNDMTcxLjYsNTEuNCwxNzMuOSw1Mi4zLDE3Ni44LDUyLjN6IE0yMDQuNywyNi4xYzAuNiwwLDEuMSwwLDEuNiwwLjFzMC45LDAuMiwxLjIsMC4zdjcuNAoJCWMtMC40LTAuMy0wLjktMC41LTEuNy0wLjhjLTAuNy0wLjMtMS42LTAuNC0yLjctMC40Yy0xLjgsMC0zLjMsMC44LTQuNSwyLjNzLTEuOSwzLjgtMS45LDd2MTUuNmgtNy4zdi0zMWg3LjN2NC45aDAuMQoJCWMwLjctMS43LDEuNy0zLDMtNEMyMDEuMiwyNi42LDIwMi44LDI2LjEsMjA0LjcsMjYuMXogTTIwNy45LDQyLjZjMC01LjEsMS40LTkuMiw0LjMtMTIuMnM2LjktNC41LDEyLjEtNC41CgkJYzQuOCwwLDguNiwxLjQsMTEuMyw0LjNjMi43LDIuOSw0LjEsNi44LDQuMSwxMS43YzAsNS0xLjQsOS00LjMsMTJjLTIuOSwzLTYuOCw0LjUtMTEuOCw0LjVjLTQuOCwwLTguNi0xLjQtMTEuNC00LjIKCQlDMjA5LjMsNTEuMywyMDcuOSw0Ny40LDIwNy45LDQyLjZ6IE0yMTUuNSw0Mi4zYzAsMy4yLDAuNyw1LjcsMi4yLDcuNGMxLjUsMS43LDMuNiwyLjYsNi4zLDIuNmMyLjcsMCw0LjctMC45LDYuMS0yLjYKCQlzMi4xLTQuMiwyLjEtNy42YzAtMy4zLTAuNy01LjgtMi4yLTcuNWMtMS40LTEuNy0zLjQtMi41LTYtMi41Yy0yLjcsMC00LjcsMC45LTYuMiwyLjdDMjE2LjIsMzYuNSwyMTUuNSwzOSwyMTUuNSw0Mi4zegoJCSBNMjUwLjUsMzQuOGMwLDEsMC4zLDEuOSwxLDIuNXMyLjEsMS4zLDQuNCwyLjJjMi45LDEuMiw1LDIuNSw2LjEsMy45YzEuMiwxLjUsMS44LDMuMiwxLjgsNS4zYzAsMi45LTEuMSw1LjMtMy40LDcKCQljLTIuMiwxLjgtNS4zLDIuNy05LjEsMi43Yy0xLjMsMC0yLjctMC4yLTQuMy0wLjVjLTEuNi0wLjMtMi45LTAuNy00LTEuMnYtNy4yYzEuMywwLjksMi44LDEuNyw0LjMsMi4yczIuOSwwLjgsNC4yLDAuOAoJCWMxLjYsMCwyLjktMC4yLDMuNi0wLjdjMC44LTAuNSwxLjItMS4yLDEuMi0yLjNjMC0xLTAuNC0xLjktMS4yLTIuNWMtMC44LTAuNy0yLjQtMS41LTQuNi0yLjRjLTIuNy0xLjEtNC42LTIuNC01LjctMy44CgkJcy0xLjctMy4yLTEuNy01LjRjMC0yLjgsMS4xLTUuMSwzLjMtNi45YzIuMi0xLjgsNS4xLTIuNyw4LjYtMi43YzEuMSwwLDIuMywwLjEsMy42LDAuNGMxLjMsMC4yLDIuNSwwLjYsMy40LDAuOVYzNAoJCWMtMS0wLjYtMi4xLTEuMi0zLjQtMS43cy0yLjYtMC43LTMuOC0wLjdjLTEuNCwwLTIuNSwwLjMtMy4yLDAuOEMyNTAuOSwzMy4xLDI1MC41LDMzLjgsMjUwLjUsMzQuOHogTTI2Ni45LDQyLjYKCQljMC01LjEsMS40LTkuMiw0LjMtMTIuMnM2LjktNC41LDEyLjEtNC41YzQuOCwwLDguNiwxLjQsMTEuMyw0LjNjMi43LDIuOSw0LjEsNi44LDQuMSwxMS43YzAsNS0xLjQsOS00LjMsMTIKCQljLTIuOSwzLTYuOCw0LjUtMTEuOCw0LjVjLTQuOCwwLTguNi0xLjQtMTEuNC00LjJDMjY4LjQsNTEuMywyNjYuOSw0Ny40LDI2Ni45LDQyLjZ6IE0yNzQuNSw0Mi4zYzAsMy4yLDAuNyw1LjcsMi4yLDcuNAoJCWMxLjUsMS43LDMuNiwyLjYsNi4zLDIuNmMyLjcsMCw0LjctMC45LDYuMS0yLjZzMi4xLTQuMiwyLjEtNy42YzAtMy4zLTAuNy01LjgtMi4yLTcuNWMtMS40LTEuNy0zLjQtMi41LTYtMi41CgkJYy0yLjcsMC00LjcsMC45LTYuMiwyLjdDMjc1LjMsMzYuNSwyNzQuNSwzOSwyNzQuNSw0Mi4zeiBNMzIyLjksMzIuNkgzMTJ2MjVoLTcuNHYtMjVoLTUuMnYtNmg1LjJ2LTQuM2MwLTMuMywxLjEtNS45LDMuMi04CgkJczQuOC0zLjEsOC4xLTMuMWMwLjksMCwxLjcsMCwyLjQsMC4xYzAuNywwLjEsMS4zLDAuMiwxLjgsMC40VjE4Yy0wLjItMC4xLTAuNy0wLjMtMS4zLTAuNWMtMC42LTAuMi0xLjMtMC4zLTIuMS0wLjMKCQljLTEuNSwwLTIuNywwLjUtMy41LDEuNGMtMC44LDEtMS4yLDIuNC0xLjIsNC4ydjMuN2gxMC45di03bDcuMy0yLjJ2OS4yaDcuNHY2aC03LjRWNDdjMCwxLjksMC4zLDMuMywxLDRjMC43LDAuOCwxLjgsMS4yLDMuMywxLjIKCQljMC40LDAsMC45LTAuMSwxLjUtMC4zYzAuNi0wLjIsMS4xLTAuNCwxLjYtMC43djZjLTAuNSwwLjMtMS4yLDAuNS0yLjMsMC43Yy0xLjEsMC4yLTIuMSwwLjMtMy4yLDAuM2MtMy4xLDAtNS40LTAuOC02LjktMi41CgkJYy0xLjUtMS42LTIuMy00LjEtMi4zLTcuNFYzMi42eiIvPgoJPGc+CgkJPHJlY3QgY2xhc3M9InN0MSIgd2lkdGg9IjM0LjIiIGhlaWdodD0iMzQuMiIvPgoJCTxyZWN0IHg9IjM3LjgiIGNsYXNzPSJzdDIiIHdpZHRoPSIzNC4yIiBoZWlnaHQ9IjM0LjIiLz4KCQk8cmVjdCB5PSIzNy44IiBjbGFzcz0ic3QzIiB3aWR0aD0iMzQuMiIgaGVpZ2h0PSIzNC4yIi8+CgkJPHJlY3QgeD0iMzcuOCIgeT0iMzcuOCIgY2xhc3M9InN0NCIgd2lkdGg9IjM0LjIiIGhlaWdodD0iMzQuMiIvPgoJPC9nPgo8L2c+Cjwvc3ZnPgo="
                                                }
                                                data-themekey="#" />
                                            &nbsp;&nbsp;<NavbarBrand tag={Link} to="/">{appTitle}</NavbarBrand>
                                        </div>) : ''}
                                
                                <div>
                                    <div className="col-auto">
                                        &nbsp;
                                    </div>
                                </div>
                                    <div className="col-auto ml-auto" style={{ textAlign: 'center', display: 'inline-block', paddingRight: '0px'}}>
                                        
                                            <div className="container">

                                            <div className="row">
                                                {showLogo ?
                                                    '' : (<div className="col" style={{ paddingLeft: '0px', paddingRight: '0px' }} >
                                                        <IconButton title='Home' iconProps={{ iconName: 'Home', styles: { root: { color: 'white' } } }} onClick={() => window.location.href = '/'} />
                                                    </div>)}

                                        {globalState.notifications.length > 0 ? (                                
                                            <div className="col" style={{ paddingLeft: '15px', paddingRight: '0px' }} >
                                        <div style={{ position: 'relative', textAlign: 'center', display: 'inline-block', paddingRight: '0px', paddingLeft: '0px' }} >
                                            <div style={{ position: 'absolute', marginLeft: '-16px', marginRight: '8px' }} >
                                            <IconButton
                                                title='Notifications'
                                                iconProps={{
                                                    iconName: globalState.notifications && globalState.notifications.length > 0 ? 'RingerSolid' : 'Ringer',
                                                    styles: { root: { color: 'white' }}
                                                }} onClick={() => this._openNotificationsPanel()} />
                                        </div>
                                            <div style={{ zIndex: 10000 }} >
                                            <CounterBadge
                                                count={globalState.notifications.length}
                                                appearance="filled"
                                                size="small"
                                                onClick={() => this._openNotificationsPanel()}
                                            />
                                            </div>
                                                </div>
                                                    </div>
                                                ) : (
                                                <div className="col" style={{ paddingLeft: '0px', paddingRight: '0px' }} >
                                                    <IconButton
                                                        title='Notifications'
                                                        iconProps={{
                                                            iconName: globalState.notifications && globalState.notifications.length > 0 ? 'RingerSolid' : 'Ringer',
                                                            styles: { root: { color: 'white' } }
                                                        }} onClick={() => this._openNotificationsPanel()} />
                                                </div>)}
                                            <div className="col" style={{ paddingLeft: '0px', paddingRight: '0px' }} >
                                            <IconButton title='Settings' iconProps={{ iconName: 'Settings', styles: { root: { color: 'white' } } }} onClick={() => this._onOpenSettings()} />             
                                            </div>
                                            </div>
                                        </div>
                                </div>
                            </div>
                        </div>

                            <Panel
                                isOpen={isMenuPanelOpen}
                                onDismiss={this._onDismisMenuPanel}
                                type={PanelType.smallFixedNear}
                                hasCloseButton={false}
                                isLightDismiss
                                onRenderHeader={(props: any, defaultRender: any) => (
                                    <div style={{ justifyContent: 'flex-start', paddingLeft: '24px' }} >
                                        <svg className="waffleMenuBlack" xmlns="http://www.w3.org/2000/svg" onClick={() => this._onDismisMenuPanel()}>
                                            <path d="M1.6 0C1.82222 0 2.03056 0.0416667 2.225 0.125C2.41944 0.208333 2.58889 0.322222 2.73333 0.466667C2.87778 0.611111 2.99167 0.780556 3.075 0.975C3.15833 1.16944 3.2 1.37778 3.2 1.6C3.2 1.82222 3.15833 2.03056 3.075 2.225C2.99167 2.41944 2.87778 2.58889 2.73333 2.73333C2.58889 2.87778 2.41944 2.99167 2.225 3.075C2.03056 3.15833 1.82222 3.2 1.6 3.2C1.37778 3.2 1.16944 3.15833 0.975 3.075C0.780556 2.99167 0.611111 2.87778 0.466667 2.73333C0.322222 2.58889 0.208333 2.41944 0.125 2.225C0.0416667 2.03056 0 1.82222 0 1.6C0 1.37778 0.0416667 1.16944 0.125 0.975C0.208333 0.780556 0.322222 0.611111 0.466667 0.466667C0.611111 0.322222 0.780556 0.208333 0.975 0.125C1.16944 0.0416667 1.37778 0 1.6 0ZM8 0C8.22222 0 8.43056 0.0416667 8.625 0.125C8.81944 0.208333 8.98889 0.322222 9.13333 0.466667C9.27778 0.611111 9.39167 0.780556 9.475 0.975C9.55833 1.16944 9.6 1.37778 9.6 1.6C9.6 1.82222 9.55833 2.03056 9.475 2.225C9.39167 2.41944 9.27778 2.58889 9.13333 2.73333C8.98889 2.87778 8.81944 2.99167 8.625 3.075C8.43056 3.15833 8.22222 3.2 8 3.2C7.77778 3.2 7.56944 3.15833 7.375 3.075C7.18056 2.99167 7.01111 2.87778 6.86667 2.73333C6.72222 2.58889 6.60833 2.41944 6.525 2.225C6.44167 2.03056 6.4 1.82222 6.4 1.6C6.4 1.37778 6.44167 1.16944 6.525 0.975C6.60833 0.780556 6.72222 0.611111 6.86667 0.466667C7.01111 0.322222 7.18056 0.208333 7.375 0.125C7.56944 0.0416667 7.77778 0 8 0ZM14.4 3.2C14.1778 3.2 13.9694 3.15833 13.775 3.075C13.5806 2.99167 13.4111 2.87778 13.2667 2.73333C13.1222 2.58889 13.0083 2.41944 12.925 2.225C12.8417 2.03056 12.8 1.82222 12.8 1.6C12.8 1.37778 12.8417 1.16944 12.925 0.975C13.0083 0.780556 13.1222 0.611111 13.2667 0.466667C13.4111 0.322222 13.5806 0.208333 13.775 0.125C13.9694 0.0416667 14.1778 0 14.4 0C14.6222 0 14.8306 0.0416667 15.025 0.125C15.2194 0.208333 15.3889 0.322222 15.5333 0.466667C15.6778 0.611111 15.7917 0.780556 15.875 0.975C15.9583 1.16944 16 1.37778 16 1.6C16 1.82222 15.9583 2.03056 15.875 2.225C15.7917 2.41944 15.6778 2.58889 15.5333 2.73333C15.3889 2.87778 15.2194 2.99167 15.025 3.075C14.8306 3.15833 14.6222 3.2 14.4 3.2ZM1.6 6.4C1.82222 6.4 2.03056 6.44167 2.225 6.525C2.41944 6.60833 2.58889 6.72222 2.73333 6.86667C2.87778 7.01111 2.99167 7.18056 3.075 7.375C3.15833 7.56944 3.2 7.77778 3.2 8C3.2 8.22222 3.15833 8.43056 3.075 8.625C2.99167 8.81944 2.87778 8.98889 2.73333 9.13333C2.58889 9.27778 2.41944 9.39167 2.225 9.475C2.03056 9.55833 1.82222 9.6 1.6 9.6C1.37778 9.6 1.16944 9.55833 0.975 9.475C0.780556 9.39167 0.611111 9.27778 0.466667 9.13333C0.322222 8.98889 0.208333 8.81944 0.125 8.625C0.0416667 8.43056 0 8.22222 0 8C0 7.77778 0.0416667 7.56944 0.125 7.375C0.208333 7.18056 0.322222 7.01111 0.466667 6.86667C0.611111 6.72222 0.780556 6.60833 0.975 6.525C1.16944 6.44167 1.37778 6.4 1.6 6.4ZM8 6.4C8.22222 6.4 8.43056 6.44167 8.625 6.525C8.81944 6.60833 8.98889 6.72222 9.13333 6.86667C9.27778 7.01111 9.39167 7.18056 9.475 7.375C9.55833 7.56944 9.6 7.77778 9.6 8C9.6 8.22222 9.55833 8.43056 9.475 8.625C9.39167 8.81944 9.27778 8.98889 9.13333 9.13333C8.98889 9.27778 8.81944 9.39167 8.625 9.475C8.43056 9.55833 8.22222 9.6 8 9.6C7.77778 9.6 7.56944 9.55833 7.375 9.475C7.18056 9.39167 7.01111 9.27778 6.86667 9.13333C6.72222 8.98889 6.60833 8.81944 6.525 8.625C6.44167 8.43056 6.4 8.22222 6.4 8C6.4 7.77778 6.44167 7.56944 6.525 7.375C6.60833 7.18056 6.72222 7.01111 6.86667 6.86667C7.01111 6.72222 7.18056 6.60833 7.375 6.525C7.56944 6.44167 7.77778 6.4 8 6.4ZM14.4 6.4C14.6222 6.4 14.8306 6.44167 15.025 6.525C15.2194 6.60833 15.3889 6.72222 15.5333 6.86667C15.6778 7.01111 15.7917 7.18056 15.875 7.375C15.9583 7.56944 16 7.77778 16 8C16 8.22222 15.9583 8.43056 15.875 8.625C15.7917 8.81944 15.6778 8.98889 15.5333 9.13333C15.3889 9.27778 15.2194 9.39167 15.025 9.475C14.8306 9.55833 14.6222 9.6 14.4 9.6C14.1778 9.6 13.9694 9.55833 13.775 9.475C13.5806 9.39167 13.4111 9.27778 13.2667 9.13333C13.1222 8.98889 13.0083 8.81944 12.925 8.625C12.8417 8.43056 12.8 8.22222 12.8 8C12.8 7.77778 12.8417 7.56944 12.925 7.375C13.0083 7.18056 13.1222 7.01111 13.2667 6.86667C13.4111 6.72222 13.5806 6.60833 13.775 6.525C13.9694 6.44167 14.1778 6.4 14.4 6.4ZM1.6 12.8C1.82222 12.8 2.03056 12.8417 2.225 12.925C2.41944 13.0083 2.58889 13.1222 2.73333 13.2667C2.87778 13.4111 2.99167 13.5806 3.075 13.775C3.15833 13.9694 3.2 14.1778 3.2 14.4C3.2 14.6222 3.15833 14.8306 3.075 15.025C2.99167 15.2194 2.87778 15.3889 2.73333 15.5333C2.58889 15.6778 2.41944 15.7917 2.225 15.875C2.03056 15.9583 1.82222 16 1.6 16C1.37778 16 1.16944 15.9583 0.975 15.875C0.780556 15.7917 0.611111 15.6778 0.466667 15.5333C0.322222 15.3889 0.208333 15.2194 0.125 15.025C0.0416667 14.8306 0 14.6222 0 14.4C0 14.1778 0.0416667 13.9694 0.125 13.775C0.208333 13.5806 0.322222 13.4111 0.466667 13.2667C0.611111 13.1222 0.780556 13.0083 0.975 12.925C1.16944 12.8417 1.37778 12.8 1.6 12.8ZM8 12.8C8.22222 12.8 8.43056 12.8417 8.625 12.925C8.81944 13.0083 8.98889 13.1222 9.13333 13.2667C9.27778 13.4111 9.39167 13.5806 9.475 13.775C9.55833 13.9694 9.6 14.1778 9.6 14.4C9.6 14.6222 9.55833 14.8306 9.475 15.025C9.39167 15.2194 9.27778 15.3889 9.13333 15.5333C8.98889 15.6778 8.81944 15.7917 8.625 15.875C8.43056 15.9583 8.22222 16 8 16C7.77778 16 7.56944 15.9583 7.375 15.875C7.18056 15.7917 7.01111 15.6778 6.86667 15.5333C6.72222 15.3889 6.60833 15.2194 6.525 15.025C6.44167 14.8306 6.4 14.6222 6.4 14.4C6.4 14.1778 6.44167 13.9694 6.525 13.775C6.60833 13.5806 6.72222 13.4111 6.86667 13.2667C7.01111 13.1222 7.18056 13.0083 7.375 12.925C7.56944 12.8417 7.77778 12.8 8 12.8ZM14.4 12.8C14.6222 12.8 14.8306 12.8417 15.025 12.925C15.2194 13.0083 15.3889 13.1222 15.5333 13.2667C15.6778 13.4111 15.7917 13.5806 15.875 13.775C15.9583 13.9694 16 14.1778 16 14.4C16 14.6222 15.9583 14.8306 15.875 15.025C15.7917 15.2194 15.6778 15.3889 15.5333 15.5333C15.3889 15.6778 15.2194 15.7917 15.025 15.875C14.8306 15.9583 14.6222 16 14.4 16C14.1778 16 13.9694 15.9583 13.775 15.875C13.5806 15.7917 13.4111 15.6778 13.2667 15.5333C13.1222 15.3889 13.0083 15.2194 12.925 15.025C12.8417 14.8306 12.8 14.6222 12.8 14.4C12.8 14.1778 12.8417 13.9694 12.925 13.775C13.0083 13.5806 13.1222 13.4111 13.2667 13.2667C13.4111 13.1222 13.5806 13.0083 13.775 12.925C13.9694 12.8417 14.1778 12.8 14.4 12.8Z"></path>
                                        </svg>
                                    </div>
                            )
                        }
                            >
                                <div className="container" style={{paddingLeft: '0px'}} >
                                    {navLinkGroups.map((nlg: INavLinkGroup) => {
                                        return (
                                            <>
                                                <div className="row" style={{ marginTop: '24px', marginBottom: '12px' }}>
                                                    <div className="col">
                                                        <Text variant='large'><b>{nlg.name}</b></Text>
                                                    </div>
                                                </div>

                                                {nlg.links.map((nl: INavLink) => {
                                                    return (
                                                        <div className="row" style={{ marginTop: '8px', marginBottom: '8px'}} >
                                                            <div className="col">
                                                                <Button
                                                                    buttonType={ButtonType.command}
                                                                    onRenderIcon={(props: IButtonProps | undefined, defaultRender: any) => {
                                                                        return (<Image src={props?.iconProps?.iconName} style={props?.iconProps?.style} />);
                                                                    }}
                                                                    href={nl.url}
                                                                    iconProps={{ iconType: IconType.image, iconName: nl.icon, style: { width: '26px', height: '26px', marginRight: '8px' } }} >
                                                                    {nl.name}
                                                                </Button>
                                                            </div>
                                                        </div>
                                                        );
                                                })}
                                            </>
                                        );
                                    })
                                    }
                                </div>
                                
                        </Panel>
                        <Notifications />
                </Container>
            </Navbar>
          </header>
        );
    }

    componentDidMount() {
        this._handleResize();

        var languages: IDropdownOption[] = [
            { key: 'en', text: 'English' },
            { key: 'de', text: 'German' }
        ];

        var branding = sessionStorage.getItem('shh_branding');
        var brandingObj: IApplicationBrandingSettings | undefined = undefined;

        if (branding !== null && branding !== undefined && branding.trim() !== '')
            try {
                brandingObj = JSON.parse(branding);
            } catch {
                brandingObj = {
                    enabled: false,
                    logo: '',
                    backgroundColor: ''
                }
            }

        var appTitle = sessionStorage.getItem('shh_appTitle');

        var navLinkGroups: INavLinkGroup[] = [];

        const seriviceHubNavLinkGroups: IServiceHubNavLinkGroup[] = [
            {
                name: 'Apps',
                expandAriaLabel: 'Expand Apps section',
                collapseAriaLabel: 'Collapse Apps section',
                links: [
                    {
                        key: 'svcM365',
                        name: 'Microsoft 365',
                        url: '/',
                        icon: '/images/m365.svg',
                        roles: ['Admin', 'ServiceHealthReader', 'Communication.Write.All']
                    },
                    {
                        key: 'svcAzure',
                        name: 'Microsoft Azure',
                        url: '/azure',
                        icon: '/images/azure.svg',
                        roles: ['Admin', 'ServiceHealthReader', 'Communication.Write.All']
                    },
                    {
                        key: 'svcPublicDashboard',
                        name: 'Public dashboard',
                        url: '/public',
                        icon: '/images/module.svg',
                        roles: ['Admin', 'ServiceHealthReader', 'Communication.Write.All', 'LicenseReader', 'Public']
                    },
                    {
                        key: 'svcAdmin',
                        name: 'Admin Center',
                        url: '/admin',
                        icon: 'data:image/svg+xml;charset=utf-8;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyMDQ4LjI1IDIwNDguMjUiPg0KICA8dGl0bGU+QWRtaW48L3RpdGxlPg0KICA8ZyBpZD0idW5pRUQ2QSI+PHBhdGggZD0iTTE0MDcgMTAyM3EwIC0yNiAtMTAgLTQ5dC0yNy41IC00MHQtNDAuNSAtMjd0LTQ5IC0xMHYyNTRxMjYgMCA0OSAtMTB0NDAuNSAtMjcuNXQyNy41IC00MXQxMCAtNDkuNXpNMjAyMCA5MzlsLTYzIC0yNDBsLTIxMCAyOHEtMTcgLTI2IC0zNS41IC01MHQtNDAuNSAtNDZsNzMgLTE5MWwtMjE0IC0xMjVsLTEyOCAxNjdxLTYwIC0xNSAtMTIyIC0xNXYyMDlxNzIgMCAxMzUgMjd0MTEwIDc0LjV0NzQgMTEwLjV0MjcgMTM1cTAgNzEgLTI3IDEzNA0KICB0LTc0IDExMHQtMTEwIDc0LjV0LTEzNSAyNy41djIxNWg1bDc5IDE3OGwyNDAgLTYzbC0yNyAtMjA0cTI3IC0xNyA1MS41IC0zN3Q0Ny41IC00M2wxODcgNzFsMTI0IC0yMTRsLTE2NiAtMTI3cTEzIC01OSAxMyAtMTI0ek02MTkgMTI2OXExNyAtODAgMzYgLTE1OC41dDM4IC0xNTguNWwtMTUxIDJxMTkgNzkgMzggMTU3LjV0MzYgMTU3LjVoM3pNNjQgMTc3NmwxMTUyIDIwNXYtMTkxNGwtMTE1MiAxOTl2MTUxMHpNMzQwIDY3N2wxMzQgLTgNCiAgbDQwIDE1OWwyMTEgLTdsNDQgLTE3MGwxNTYgLTEwbC0yMjQgNzY2bC0xNTggLTEweiIgZmlsbD0iIzU4NTk1QiIgdHJhbnNmb3JtPSJzY2FsZSgxLC0xKSB0cmFuc2xhdGUoMCwgLTIwNDguMjUpIiAvPjwvZz4NCjwvc3ZnPg==',
                        roles: ['Admin']
                    }
                ],
            }
        ];

        acquireAccessToken()
            .then((result) => {
                var tokenClaims: any = result.account?.idTokenClaims;
                const roles: any = tokenClaims?.roles;

                if (roles !== undefined) {
                    for (const navGroups of seriviceHubNavLinkGroups) {
                        var navLinks: IServiceHubNavLink[] = [];
                        for (const link of navGroups.links) {
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
                                    url: nLink.url,
                                    icon: nLink.icon
                                });
                            }

                            navLinkGroups.push({
                                name: navGroups.name,
                                expandAriaLabel: navGroups.expandAriaLabel,
                                collapseAriaLabel: navGroups.collapseAriaLabel,
                                links: ngLinks
                            });
                        }
                    }
                }

                this.setState({
                    navLinkGroups: navLinkGroups,
                    languageOptions: languages,
                    branding: brandingObj!,
                    appTitle: appTitle && appTitle !== null ? appTitle : 'Service Health Hub'
                });
            });
    }

    private _onOpenMenuPanel = (): void => {
        this.setState({
            isMenuPanelOpen: true
        });
    }

    private _onDismisMenuPanel = (): void => {
        this.setState({
            isMenuPanelOpen: false
        });
    }

    private _openNotificationsPanel = (): void => {
        let globalState = this.context;
        globalState.openNotificationsPanel(true);
    }

    private _onOpenSettings = (): void => {
        window.location.href = "/settings";
    }

    _handleResize = () => {
        var clientWidth: number = window.innerWidth;

        this.setState({
            showLogo: clientWidth < 500 ? false : true,
        });
    }

    private _getWidth() {
        return Math.max(
            document.body.scrollWidth,
            document.documentElement.scrollWidth,
            document.body.offsetWidth,
            document.documentElement.offsetWidth,
            document.documentElement.clientWidth
        );
    }

    private _getHeight() {
        return Math.max(
            document.body.scrollHeight,
            document.documentElement.scrollHeight,
            document.body.offsetHeight,
            document.documentElement.offsetHeight,
            document.documentElement.clientHeight
        );
    }
}
