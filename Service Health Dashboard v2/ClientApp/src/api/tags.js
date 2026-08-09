"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.getAvailableTagDefinitions = exports.getTags = exports.removeTag = exports.addTag = void 0;
var AccessTokenHelper_1 = require("../auth/AccessTokenHelper");
function addTag(messageId, type, tagId, onSuccess, onError) {
    var params;
    AccessTokenHelper_1.acquireAccessToken()
        .then(function (response) {
        var body = {
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
        var responseObj = {
            ok: false,
            statusCode: null,
            body: null
        };
        fetch("/api/tag/" + type + "/" + messageId, params)
            .then(function (response) {
            responseObj.ok = response.ok;
            responseObj.statusCode = response.status;
            return response.ok ? null : response.text();
        })
            .then(function (body) {
            responseObj.body = body;
            if (!responseObj.ok) {
                // make the promise be rejected if we didn't get a 2xx response
                var err = new Error("Couldn't add tag. Error details: HTTP " + responseObj.statusCode + " - " + body);
                throw err;
            }
            if (onSuccess)
                onSuccess(messageId, type, tagId);
        })
            .catch(function (err) {
            if (onError)
                onError(err.message);
        });
    })
        .catch(function (err) {
        if (onError)
            onError("Authentication error. Details: " + err.message);
    });
    ;
}
exports.addTag = addTag;
function removeTag(messageId, type, tagId, onSuccess, onError) {
    var params;
    AccessTokenHelper_1.acquireAccessToken()
        .then(function (response) {
        params = {
            headers: {
                "Content-Type": "application/json charset=UTF-8",
                "Authorization": "Bearer " + response.idToken
            },
            method: "DELETE"
        };
        var responseObj = {
            ok: false,
            statusCode: null,
            body: null
        };
        fetch("/api/tag/" + type + "/" + messageId + "/" + tagId, params)
            .then(function (response) {
            responseObj.ok = response.ok;
            responseObj.statusCode = response.status;
            return response.ok ? null : response.text();
        })
            .then(function (body) {
            responseObj.body = body;
            if (!responseObj.ok) {
                // make the promise be rejected if we didn't get a 2xx response
                var err = new Error("Couldn't delete tag. Error details: HTTP " + responseObj.statusCode + " - " + body);
                throw err;
            }
            if (onSuccess)
                onSuccess(messageId, type, tagId);
        })
            .catch(function (err) {
            if (onError)
                onError(err.message);
        });
    })
        .catch(function (err) {
        if (onError)
            onError("Authentication error. Details: " + err.message);
    });
    ;
}
exports.removeTag = removeTag;
function getTags(messageId, type, onSuccess, onError) {
    var params;
    AccessTokenHelper_1.acquireAccessToken()
        .then(function (response) {
        params = {
            headers: {
                "Content-Type": "application/json charset=UTF-8",
                "Authorization": "Bearer " + response.idToken
            },
            method: "GET"
        };
        var responseObj = {
            ok: false,
            statusCode: null,
            body: null
        };
        fetch("/api/tag/" + type + "/" + messageId, params)
            .then(function (response) {
            responseObj.ok = response.ok;
            responseObj.statusCode = response.status;
            return response.ok ? response.json() : response.text();
        })
            .then(function (body) {
            if (!responseObj.ok) {
                // make the promise be rejected if we didn't get a 2xx response
                var err = new Error("Couldn't retrieve tags. Error details: HTTP " + responseObj.statusCode + " - " + body);
                throw err;
            }
            var res = [];
            for (var _i = 0, body_1 = body; _i < body_1.length; _i++) {
                var tag = body_1[_i];
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
            .catch(function (err) {
            if (onError)
                onError(err.message);
        });
    })
        .catch(function (err) {
        if (onError)
            onError("Authentication error. Details: " + err.message);
    });
    ;
}
exports.getTags = getTags;
function getAvailableTagDefinitions(type, onSuccess, onError) {
    var params;
    AccessTokenHelper_1.acquireAccessToken()
        .then(function (response) {
        params = {
            headers: {
                "Content-Type": "application/json charset=UTF-8",
                "Authorization": "Bearer " + response.idToken
            },
            method: "GET"
        };
        var responseObj = {
            ok: false,
            statusCode: null,
            body: null
        };
        fetch("/api/tagdefinitions?componentName=" + type, params)
            .then(function (response) {
            responseObj.ok = response.ok;
            responseObj.statusCode = response.status;
            return response.ok ? response.json() : response.text();
        })
            .then(function (body) {
            if (!responseObj.ok) {
                // make the promise be rejected if we didn't get a 2xx response
                var err = new Error("Couldn't retrieve tags. Error details: HTTP " + responseObj.statusCode + " - " + body);
                throw err;
            }
            var items = [];
            for (var _i = 0, body_2 = body; _i < body_2.length; _i++) {
                var groupedTagDefinition = body_2[_i];
                if (groupedTagDefinition.tagDefinitions !== null && groupedTagDefinition.tagDefinitions !== undefined) {
                    for (var _a = 0, _b = groupedTagDefinition.tagDefinitions; _a < _b.length; _a++) {
                        var tagDefinition = _b[_a];
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
            }
            items = items.sort(function (a, b) { return a.name > b.name ? -1 : 1; });
            if (onSuccess)
                onSuccess(items);
        })
            .catch(function (err) {
            if (onError)
                onError(err.message);
        });
    })
        .catch(function (err) {
        if (onError)
            onError("Authentication error. Details: " + err.message);
    });
    ;
}
exports.getAvailableTagDefinitions = getAvailableTagDefinitions;
//# sourceMappingURL=tags.js.map