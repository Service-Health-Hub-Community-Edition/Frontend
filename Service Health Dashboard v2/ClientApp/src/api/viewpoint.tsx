import { acquireAccessToken } from '../auth/AccessTokenHelper';

export function setViewState(id: string, state: boolean, onSuccess?: (id: string, state: boolean) => void, onError?: (message: string) => void): void {
    var params: RequestInit;
    acquireAccessToken()
            .then((response) => {
        const body = {
            communicationId: id,
            read: state
        };

        params = {
            headers: {
                "Content-Type": "application/json charset=UTF-8",
                "Authorization": "Bearer " + response.idToken
            },
            body: JSON.stringify(body),
            method: "POST"
        };

        var responseObj: any = {
            ok: false,
            statusCode: null,
            body: null
        };

        fetch("/api/items/viewpoint", params)
            .then((response) => {
                responseObj.ok = response.ok;
                responseObj.statusCode = response.status;
                return response.ok ? null : response.text();
            })
            .then((body) => {
                responseObj.body = body;

                if (!responseObj.ok) {
                    // make the promise be rejected if we didn't get a 2xx response
                    const err = new Error("Couldn't set viewpoint state. Error details:<br/><br/><b>HTTP " + responseObj.statusCode + "</b><br/>" + body);
                    throw err;
                }

                if (onSuccess)
                    onSuccess(id, state);
            })
            .catch((err) => {
                if (onError)
                    onError(err.message);
            });
    })
        .catch((err) => {
            if (onError)
                onError("Authentication error. Details:<br/><br/>" + err.message);
        });;
}

export function setArchiveState(id: string, state: boolean, onSuccess?: (id: string, state: boolean) => void, onError?: (message: string) => void): void {
    var params: RequestInit;
    acquireAccessToken()
        .then((response) => {
            const body = {
                communicationId: id,
                archived: state
            };

            params = {
                headers: {
                    "Content-Type": "application/json charset=UTF-8",
                    "Authorization": "Bearer " + response.idToken
                },
                body: JSON.stringify(body),
                method: "POST"
            };

            var responseObj: any = {
                ok: false,
                statusCode: null,
                body: null
            };

            fetch("/api/items/viewpoint", params)
                .then((response) => {
                    responseObj.ok = response.ok;
                    responseObj.statusCode = response.status;
                    return response.ok ? null : response.text();
                })
                .then((body) => {
                    responseObj.body = body;

                    if (!responseObj.ok) {
                        // make the promise be rejected if we didn't get a 2xx response
                        const err = new Error("Couldn't set viewpoint state. Error details:<br/><br/><b>HTTP " + responseObj.statusCode + "</b><br/>" + body);
                        throw err;
                    }

                    if (onSuccess)
                        onSuccess(id, state);
                })
                .catch((err) => {
                    if (onError)
                        onError(err.message);
                });
        })
        .catch((err) => {
            if (onError)
                onError("Authentication error. Details:<br/><br/>" + err.message);
        });;
}

export function setFavoriteState(id: string, state: boolean, onSuccess?: (id: string, state: boolean) => void, onError?: (message: string) => void): void {
    var params: RequestInit;
    acquireAccessToken()
        .then((response) => {
            const body = {
                communicationId: id,
                favorite: state
            };

            params = {
                headers: {
                    "Content-Type": "application/json charset=UTF-8",
                    "Authorization": "Bearer " + response.idToken
                },
                body: JSON.stringify(body),
                method: "POST"
            };

            var responseObj: any = {
                ok: false,
                statusCode: null,
                body: null
            };

            fetch("/api/items/viewpoint", params)
                .then((response) => {
                    responseObj.ok = response.ok;
                    responseObj.statusCode = response.status;
                    return response.ok ? null : response.text();
                })
                .then((body) => {
                    responseObj.body = body;

                    if (!responseObj.ok) {
                        // make the promise be rejected if we didn't get a 2xx response
                        const err = new Error("Couldn't set viewpoint state. Error details:<br/><br/><b>HTTP " + responseObj.statusCode + "</b><br/>" + body);
                        throw err;
                    }

                    if (onSuccess)
                        onSuccess(id, state);
                })
                .catch((err) => {
                    if (onError)
                        onError(err.message);
                });
        })
        .catch((err) => {
            if (onError)
                onError("Authentication error. Details:<br/><br/>" + err.message);
        });;
}