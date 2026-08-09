import { AuthenticationResult } from "@azure/msal-common";
import { acquireAccessToken } from "../auth/AccessTokenHelper";

export async function translate (language: string, data: Map<string, string>, authResponse: AuthenticationResult): Promise<Map<string, string>> {
    var params: RequestInit;
    let result = new Map<string, string>();

    const body = {
        language: language,
        contents: Array.from(data.keys())
    };

    params = {
        headers: {
            "Content-Type": "application/json charset=UTF-8",
            "Authorization": "Bearer " + authResponse.idToken
        },
        body: JSON.stringify(body),
        method: "POST"
    };

    const response = await fetch("/api/Translate", params);
    const jsonResponse = await response.json();

    for (var value in jsonResponse) {
        result.set(value, jsonResponse[value])
    }

    return result;
}

export function getUIString(uiStrings: Map<string, string>, key: string): string {
    if (uiStrings !== undefined) {
        var value: string | undefined = uiStrings.get(key);

        if (value === undefined) {
            return key;
        } else {
            return value!;
        }
    } else {
        return key;
    }
}

/* export const translate = async (language: string, data: Map<string, string>): Promise<Map<string, string>> => {
    var params: RequestInit;
    let result = new Map<string, string>();

    acquireAccessToken()
        .then((authResponse) => {
            const body = {
                language: language,
                contents: Array.from(data.keys())
            };

            params = {
                headers: {
                    "Content-Type": "application/json charset=UTF-8",
                    "Authorization": "Bearer " + authResponse.idToken
                },
                body: JSON.stringify(body),
                method: "POST"
            };
        }).then(async () => {
            const response = await fetch("/api/Translate", params);
            const jsonResponse = await response.json();
         
            for (var value in jsonResponse) {
                result.set(value, jsonResponse[value])
            }
        });

    return result;
} */