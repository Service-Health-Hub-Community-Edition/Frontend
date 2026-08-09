import * as React from 'react';
import { Text } from '@fluentui/react';
import { AccessDenied } from './AccessDenied';
import { acquireAccessToken } from '../auth/AccessTokenHelper';

interface IAccessGateState {
    requiredRoles: string[];
    accessGranted: boolean | null;
}

export class AccessGate extends React.Component<{ requiredRoles: string[] }, IAccessGateState> {
    constructor(props: { requiredRoles: string[] }) {
        super(props);

        this.state = {
            requiredRoles: props.requiredRoles,
            accessGranted: null
        };
    }

    public render() {
        const { accessGranted } = this.state;

        if (accessGranted === null)
            return (<></>);

        if (!accessGranted)
            return (<AccessDenied />);

        return (
            <>
                {this.props.children}
            </>
        )
    }

    componentDidMount() {
        acquireAccessToken()
            .then((response) => {
                var tokenClaims: any = response.account?.idTokenClaims;
                const userRoles: any = tokenClaims?.roles;
                var userHasRequiredRole: boolean = userRoles.some((r: string) => this.props.requiredRoles.includes(r));

                this.setState({
                    accessGranted: userHasRequiredRole
                });
            })
            .catch(() => {
                this.setState({
                    accessGranted: false
                });
            });
    }
}