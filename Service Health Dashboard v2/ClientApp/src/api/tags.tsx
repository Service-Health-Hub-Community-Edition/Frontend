import { acquireAccessToken } from '../auth/AccessTokenHelper';
import { ITagDefinition } from '../components/admin/applicationSettings/TagDefinitions';

export interface MSTag
{
    id: number;
    messageId: string;
    type: string;
    tagId: string;
    modified?: Date | undefined | null;
}

export function addTag(messageId: string, type: string, tagId: string, onSuccess?: (messageId: string, type: string, tagId: string) => void, onError?: (message: string) => void): void {
    var params: RequestInit;
    acquireAccessToken()
            .then((response) => {
        const body = {
            op: "add",
            tagId: tagId
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

        fetch("/api/tag/" + type + "/" + messageId, params)
            .then((response) => {
                responseObj.ok = response.ok;
                responseObj.statusCode = response.status;
                return response.ok ? null : response.text();
            })
            .then((body) => {
                responseObj.body = body;

                if (!responseObj.ok) {
                    // make the promise be rejected if we didn't get a 2xx response
                    const err = new Error("Couldn't add tag. Error details: HTTP " + responseObj.statusCode + " - " + body);
                    throw err;
                }

                if (onSuccess)
                    onSuccess(messageId, type, tagId);
            })
            .catch((err) => {
                if (onError)
                    onError(err.message);
            });
    })
        .catch((err) => {
            if (onError)
                onError("Authentication error. Details: " + err.message);
        });;
}

export function removeTag(messageId: string, type: string, tagId: string, onSuccess?: (messageId: string, type: string, tagId: string) => void, onError?: (message: string) => void): void {
    var params: RequestInit;
    acquireAccessToken()
        .then((response) => {
            params = {
                headers: {
                    "Content-Type": "application/json charset=UTF-8",
                    "Authorization": "Bearer " + response.idToken
                },
                method: "DELETE"
            };

            var responseObj: any = {
                ok: false,
                statusCode: null,
                body: null
            };

            fetch("/api/tag/" + type + "/" + messageId + "/" + tagId, params)
                .then((response) => {
                    responseObj.ok = response.ok;
                    responseObj.statusCode = response.status;
                    return response.ok ? null : response.text();
                })
                .then((body) => {
                    responseObj.body = body;

                    if (!responseObj.ok) {
                        // make the promise be rejected if we didn't get a 2xx response
                        const err = new Error("Couldn't delete tag. Error details: HTTP " + responseObj.statusCode + " - " + body);
                        throw err;
                    }

                    if (onSuccess)
                        onSuccess(messageId, type, tagId);
                })
                .catch((err) => {
                    if (onError)
                        onError(err.message);
                });
        })
        .catch((err) => {
            if (onError)
                onError("Authentication error. Details: " + err.message);
        });;
}

export function getTags(messageId: string, type: string, onSuccess?: (res: MSTag[]) => void, onError?: (message: string) => void): void {
    var params: RequestInit;
    acquireAccessToken()
        .then((response) => {

            params = {
                headers: {
                    "Content-Type": "application/json charset=UTF-8",
                    "Authorization": "Bearer " + response.idToken
                },
                method: "GET"
            };

            var responseObj: any = {
                ok: false,
                statusCode: null,
                body: null
            };

            fetch("/api/tag/" + type + "/" + messageId, params)
                .then((response) => {
                    responseObj.ok = response.ok;
                    responseObj.statusCode = response.status;
                    return response.ok ? response.json() : response.text();
                })
                .then((body) => {
                    if (!responseObj.ok) {
                        // make the promise be rejected if we didn't get a 2xx response
                        const err = new Error("Couldn't retrieve tags. Error details: HTTP " + responseObj.statusCode + " - " + body);
                        throw err;
                    }

                    var res: MSTag[] = [];

                    for (const tag of body) {
                        res.push({
                            id: tag.id,
                            messageId: tag.messageId,
                            tagId: tag.tagId,
                            type: tag.type,
                            modified: tag.modified
                        });
                    }

                    if (onSuccess)
                        onSuccess(res);
                })
                .catch((err) => {
                    if (onError)
                        onError(err.message);
                });
        })
        .catch((err) => {
            if (onError)
                onError("Authentication error. Details: " + err.message);
        });;
}

export function getAvailableTagDefinitions(type: string, onSuccess?: (res: ITagDefinition[]) => void, onError?: (message: string) => void): void {
    var params: RequestInit;
    acquireAccessToken()
        .then((response) => {

            params = {
                headers: {
                    "Content-Type": "application/json charset=UTF-8",
                    "Authorization": "Bearer " + response.idToken
                },
                method: "GET"
            };

            var responseObj: any = {
                ok: false,
                statusCode: null,
                body: null
            };

            fetch("/api/tagdefinitions?componentName=" + type, params)
                .then((response) => {
                    responseObj.ok = response.ok;
                    responseObj.statusCode = response.status;
                    return response.ok ? response.json() : response.text();
                })
                .then((body) => {
                    if (!responseObj.ok) {
                        // make the promise be rejected if we didn't get a 2xx response
                        const err = new Error("Couldn't retrieve tags. Error details: HTTP " + responseObj.statusCode + " - " + body);
                        throw err;
                    }

                    var items: ITagDefinition[] = [];

                    for (const groupedTagDefinition of body) {
                        if (groupedTagDefinition.tagDefinitions !== null && groupedTagDefinition.tagDefinitions !== undefined) {
                            for (const tagDefinition of groupedTagDefinition.tagDefinitions)
                                items.push({
                                    key: tagDefinition.tagId,
                                    id: tagDefinition.id,
                                    name: tagDefinition.name,
                                    tagId: tagDefinition.tagId,
                                    type: tagDefinition.type && tagDefinition.type !== null ? tagDefinition.type : "!!!general",
                                    lastUsed: tagDefinition.lastUsed,
                                    itemCount: tagDefinition.itemCount
                                });
                        }
                    }

                    items = items.sort((a, b) => a.name > b.name ? -1 : 1);                   

                    if (onSuccess)
                        onSuccess(items);
                })
                .catch((err) => {
                    if (onError)
                        onError(err.message);
                });
        })
        .catch((err) => {
            if (onError)
                onError("Authentication error. Details: " + err.message);
        });;
}