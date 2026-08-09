import * as React from 'react';
import { GlobalState } from './GlobalState';
import { acquireAccessToken } from "../auth/AccessTokenHelper";

export async function getUserProfileAsync(): Promise<any> {
    const authResponse = await acquireAccessToken();
    const userProfileResponse = await fetch('/api/users/' + authResponse.uniqueId, { headers: authResponse.idToken === "" ? {} : { 'Authorization': `Bearer ${authResponse.idToken}` } });
    const userProfileData = userProfileResponse.ok ? await userProfileResponse.json() : await userProfileResponse.text();

    return {
        ok: userProfileResponse.ok,
        statusCode: userProfileResponse.status,
        data: userProfileData
    };
}

export async function setPropertyAsync(name: string, value: any): Promise<any> {
    const authResponse = await acquireAccessToken();

    const body = {
        [name]: value
    };

    const params = {
        headers: {
            "Content-Type": "application/json charset=UTF-8",
            "Authorization": "Bearer " + authResponse.idToken
        },
        body: JSON.stringify(body),
        method: "POST"
    };

    
    const userProfileResponse = await fetch('/api/users/' + authResponse.uniqueId, params);
    var userProfileData = userProfileResponse.ok ? await userProfileResponse.json() : await userProfileResponse.text();

    if (!userProfileResponse.ok)
        userProfileData = "Couldn't save user profile property " + name + ". Status code: " + userProfileResponse.status + ". Response: " + userProfileData;

    return {
        ok: userProfileResponse.ok,
        statusCode: userProfileResponse.status,
        data: userProfileData
    };
}

export interface IUserProfileState
{
    initialized: boolean;
}

export class UserProfile extends React.Component<{}, IUserProfileState> {
    static contextType = GlobalState;

    constructor(props: {}) {
        super(props);

        this.state = {
            initialized: false
        };
    }

    componentDidMount() {
        let globalState: any = this.context;
        getUserProfileAsync()
            .then((result: any) => {
                if (result.ok)
                    globalState.setProfile(result.data.properties ? result.data.properties : { } );
                else
                    globalState.setProfile({});

                this.setState({
                    initialized: true
                });
            })
    }

    render(): React.ReactNode {
        const { initialized } = this.state;

        return(
            <div>
                { initialized ? this.props.children : '' }
            </div>
        )
    }
}