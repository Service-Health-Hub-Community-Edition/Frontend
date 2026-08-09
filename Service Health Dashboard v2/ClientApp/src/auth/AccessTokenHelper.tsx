import { PublicClientApplication, AuthenticationResult } from "@azure/msal-browser";
import { fetchMsalConfig } from "./authConfig";

var msalInstance: PublicClientApplication;

export const acquireAccessToken = async () => {
    //Getting the configuration from the server

    if (msalInstance === undefined) {
        const msalConfig = await fetchMsalConfig();
        msalInstance = new PublicClientApplication(msalConfig);
    }

    const activeAccount = msalInstance.getActiveAccount(); // This will only return a non-null value if you have logic somewhere else that calls the setActiveAccount API
    const accounts = msalInstance.getAllAccounts();

    if (!activeAccount && accounts.length === 0) {
        /*
        * User is not signed in. Throw error or wait for user to login.
        * Do not attempt to log a user in outside of the context of MsalProvider
        */
    }
    const request = {
        scopes: ["User.Read", "email", "openid", "profile", "offline_access"],
        account: activeAccount || accounts[0]
    };

    const authResult = await msalInstance.acquireTokenSilent(request);
    return authResult;
};

export const accessControl = (token: AuthenticationResult, role: string) => {
    var tokenClaims: any = token.account?.idTokenClaims;
    if (!tokenClaims?.roles.includes(role)) {
        window.open("/AccessDenied", "_self");
    }
}
