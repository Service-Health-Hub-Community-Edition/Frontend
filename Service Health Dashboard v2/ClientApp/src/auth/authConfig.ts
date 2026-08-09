/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import { Configuration, LogLevel } from "@azure/msal-browser";
import { BrowserAuthOptions } from "@azure/msal-browser";

/**
 * Configuration object to be passed to MSAL instance on creation. 
 * For a full list of MSAL.js configuration parameters, visit:
 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/configuration.md 
 */
/* var configCache: BrowserAuthOptions; */

export const msalConfig: Configuration = {
    auth: {
        clientId: '', 
        authority: '',
        redirectUri: '/'
    },
    cache: {
        cacheLocation: "sessionStorage", // This configures where your cache will be stored
        storeAuthStateInCookie: true, // Set this to "true" if you are having issues on IE11 or Edge
    },
    system: {	
        loggerOptions: {	
            loggerCallback: (level: any, message: any, containsPii: any) => {	
                if (containsPii) {		
                    return;		
                }		
                switch (level) {		
                    case LogLevel.Error:		
                        console.error(message);		
                        return;		
                    case LogLevel.Info:		
                        console.info(message);		
                        return;		
                    case LogLevel.Verbose:		
                        console.debug(message);		
                        return;		
                    case LogLevel.Warning:		
                        console.warn(message);		
                        return;		
                }	
            }	
        }	
    }
};

/**
 * Scopes you add here will be prompted for user consent during sign-in.
 * By default, MSAL.js will add OIDC scopes (openid, profile, email) to any login request.
 * For more information about OIDC scopes, visit: 
 * https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-permissions-and-consent#openid-connect-scopes
 */
export const loginRequest = {
    scopes: ["User.Read", "email", "openid", "profile", "offline_access"]
};

/**
 * Add here the scopes to request when obtaining an access token for MS Graph API. For more information, see:
 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/resources-and-scopes.md
 */
export const graphConfig = {
    graphMeEndpoint: "https://graph.microsoft.com/v1.0/me"
};

export const fetchMsalConfig = async (): Promise<Configuration> => {
    var clientConfigOptions: string | null = sessionStorage.getItem('clientConfigOptions');
    if (clientConfigOptions === null|| clientConfigOptions === "" || clientConfigOptions === undefined) {
        const response = await fetch('/api/config');
        const jsonResponse = await response.json();
        sessionStorage.setItem("shh_dbVersion", jsonResponse.dbVersion);
        sessionStorage.setItem("shh_branding", jsonResponse.branding ? JSON.stringify(jsonResponse.branding) : "");
        sessionStorage.setItem("shh_appTitle", jsonResponse.appTitle ? jsonResponse.appTitle : "");

        const authOptions: BrowserAuthOptions = {
            clientId: jsonResponse.clientId,
            authority: "https://login.microsoftonline.com/" + jsonResponse.tenantId,
            redirectUri: '/'
        }
        sessionStorage.setItem('clientConfigOptions', JSON.stringify(authOptions));
        msalConfig.auth = authOptions;
    }
    else {
        var configCache: BrowserAuthOptions = JSON.parse(clientConfigOptions);
        msalConfig.auth = configCache;
    }
    return msalConfig;
}