import * as React from 'react';
import { Component } from 'react';
import { Redirect } from 'react-router-dom';
import { Navigate, Route, useLocation } from 'react-router-dom';
import { acquireAccessToken } from '../../auth/AccessTokenHelper';
import { AccessDenied } from '../AccessDenied';

/* const PrivateRoute = ({ children, roles, redirectUri }:
    {
        children: JSX.Element; roles: string[]; redirectUri: string | undefined
    }) => {
    console.log("in PrivateRoute");

    let location = useLocation();

    var userHasRequiredRole: boolean = false;

    console.log("in PrivateRoute");

    acquireAccessToken()
        .then((result) => {
            var tokenClaims: any = result.account?.idTokenClaims;
            const userRoles: any = tokenClaims?.roles;
            userHasRequiredRole = userRoles.some((r: string) => roles.includes(r));
        });

    if (!userHasRequiredRole) {
        if (redirectUri !== undefined && redirectUri.trim() !== "") {
            return (<Navigate to={{ redirectUri }} state={{ from: location }} />);
        } else {
            return (<Navigate to="/AccessDenied" state={{ from: location }} />);
        }
    }

    return children;
};

const PrivateRoute2 = ({ children, roles, redirectUri }:
    {
        children: JSX.Element; roles: string[]; redirectUri: string | undefined
    }) => {

    return children;
};

export default PrivateRoute; */

interface IPrivateRouteState {
    roles: string[];
    redirectUri?: string;
}

export default class PrivateRoute extends React.Component<{ roles: string[]; redirectUri?: string }, IPrivateRouteState> {
    constructor(props: { roles: string[]; redirectUri?: string }) {
        super(props);

        this.state = {
            roles: this.props.roles,
            redirectUri: this.props.redirectUri
        };
    }

    public render() {
        const {
            roles, redirectUri
        } = this.state;

        var userHasRequiredRole: boolean = false;

        acquireAccessToken()
            .then((result) => {
                var tokenClaims: any = result.account?.idTokenClaims;
                const userRoles: any = tokenClaims?.roles;
                userHasRequiredRole = userRoles.some((r: string) => roles.includes(r));
            });
        
        if (!userHasRequiredRole) {
            if (redirectUri !== undefined && redirectUri.trim() !== "") {
                return (<Navigate to={redirectUri} />);
            } else {
                return (<AccessDenied />);
            }
        }

        return this.props.children;
    }
}