import * as React from 'react';
import * as msTeams from '@microsoft/teams-js';
import { useMsal } from "@azure/msal-react";
import { loginRequest } from '../../auth/authConfig';
import { RedirectRequest } from '@azure/msal-browser';


function getHashParameters() {
    let hashParams: any = {};
    window.location.hash.substr(1).split("&").forEach(function (item) {
      let s = item.split("="),
        k = s[0],
        v = s[1] && decodeURIComponent(s[1]);
      hashParams[k] = v as any;
    });
    return hashParams;
  }


const AuthEnd: React.FC = () => {
    
    let hashParams = getHashParameters();
    /* if (hashParams["error"]) {
        // Authentication/authorization failed
        msTeams.authentication.notifyFailure(hashParams["error"]);
    } else if (hashParams["access_token"] || hashParams["code"]) {
        // Success: return token information to the tab
        msTeams.authentication.notifySuccess();
    } else {
        // Unexpected condition: hash does not contain error or access_token parameter
        msTeams.authentication.notifyFailure(JSON.stringify(hashParams));
    } */

    if (hashParams["error"]) {
        // Authentication/authorization failed
        msTeams.authentication.notifyFailure(hashParams["error"]);
    } else {
        // Success: return token information to the tab
        msTeams.authentication.notifySuccess();
    }

    const { instance, accounts, inProgress } = useMsal();

    return (
        <h1>Authenticated in Teams </h1>
    );

}
export default AuthEnd;