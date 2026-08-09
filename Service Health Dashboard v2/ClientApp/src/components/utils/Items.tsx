import * as React from 'react';
import { acquireAccessToken } from "../../auth/AccessTokenHelper";

export const setArchiveFlag = (id: string, type: string, archive: boolean): void => {
    var params: RequestInit;
    acquireAccessToken()
        .then((response) => {
            const body = {
                id: id,
                type: type,
            };

            params = {
                headers: {
                    "Content-Type": "application/json charset=UTF-8",
                    "Authorization": "Bearer " + response.idToken
                },
                body: JSON.stringify(body),
                method: "POST"
            };

            const uri: string = archive ? '/api/Items/Archive' : '/api/Items/Restore';

            fetch(uri, params)
                .then((response) => {

                    if (!response.ok) {
                        // make the promise be rejected if we didn't get a 2xx response
                        const err = new Error("Couldn't publish message. Error details:<br/><br/><b>HTTP " + response.status + "</b><br/>" + response.statusText);
                        throw err;
                    }
                })
        });
}