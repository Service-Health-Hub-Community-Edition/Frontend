"use strict";
var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
exports.JobStatistics = void 0;
var React = require("react");
var react_1 = require("@fluentui/react");
var react_charting_1 = require("@fluentui/react-charting");
var AccessTokenHelper_1 = require("../../../auth/AccessTokenHelper");
var calloutItemStyle = react_1.mergeStyles({
    borderBottom: '1px solid #D9D9D9',
    padding: '3px',
});
var JobStatistics = /** @class */ (function (_super) {
    __extends(JobStatistics, _super);
    function JobStatistics(props) {
        var _this = _super.call(this, props) || this;
        _this.state = {
            componentId: props.componentId,
            items: [],
            accessGranted: false,
            initialized: false
        };
        return _this;
    }
    JobStatistics.prototype.render = function () {
        var _a, _b;
        var _c = this.state, items = _c.items, accessGranted = _c.accessGranted, initialized = _c.initialized, error = _c.error;
        if (!initialized)
            return (React.createElement(React.Fragment, null,
                React.createElement("br", null),
                React.createElement("br", null),
                React.createElement(react_1.Spinner, { size: react_1.SpinnerSize.medium })));
        if (error !== undefined)
            return (React.createElement(React.Fragment, null,
                React.createElement("br", null),
                React.createElement("br", null),
                React.createElement(react_1.Text, { variant: 'medium' },
                    React.createElement("b", null, error))));
        if (!accessGranted)
            return (React.createElement(React.Fragment, null,
                React.createElement("br", null),
                React.createElement("br", null),
                React.createElement(react_1.Text, { variant: 'medium' },
                    React.createElement("b", null, "You have no access to this ressource."))));
        if (!items || items.length <= 0)
            return (React.createElement(React.Fragment, null,
                React.createElement("br", null),
                React.createElement("br", null),
                React.createElement(react_1.Text, { variant: 'medium' },
                    React.createElement("b", null, "No events in past 24 hours."))));
        var data = [
            {
                chartTitle: 'Completed',
                chartData: [
                    {
                        legend: 'Completed jobs',
                        horizontalBarChartdata: { x: (_a = items.filter(function (i) { return i.state === 'Completed'; })) === null || _a === void 0 ? void 0 : _a.length, y: items.length },
                        color: react_1.DefaultPalette.green,
                    },
                ],
            },
            {
                chartTitle: 'Failed',
                chartData: [
                    {
                        legend: 'Failed jobs',
                        horizontalBarChartdata: { x: (_b = items.filter(function (i) { return i.state === 'Failed'; })) === null || _b === void 0 ? void 0 : _b.length, y: items.length },
                        color: react_1.DefaultPalette.redDark,
                    },
                ],
            }
        ];
        return (React.createElement(React.Fragment, null,
            React.createElement(react_charting_1.HorizontalBarChart, { culture: window.navigator.language, data: data, hideRatio: [true] })));
    };
    JobStatistics.prototype.componentDidMount = function () {
        this.setState({
            componentId: this.props.componentId
        });
        this._loadData();
    };
    JobStatistics.prototype._loadData = function () {
        var _this = this;
        var requiredRoles = ['Admin'];
        var authResponse;
        var userHasRequiredRole = false;
        this.setState({
            initialized: false
        });
        AccessTokenHelper_1.acquireAccessToken()
            .then(function (response) {
            var _a;
            var tokenClaims = (_a = response.account) === null || _a === void 0 ? void 0 : _a.idTokenClaims;
            var userRoles = tokenClaims === null || tokenClaims === void 0 ? void 0 : tokenClaims.roles;
            userHasRequiredRole = userRoles.some(function (r) { return requiredRoles.includes(r); });
            _this.setState({
                accessGranted: userHasRequiredRole
            });
            authResponse = response;
        })
            .then(function () {
            if (userHasRequiredRole)
                fetch('/api/jobs/' + _this.state.componentId + '/history', { headers: authResponse.idToken === "" ? {} : { 'Authorization': "Bearer " + authResponse.idToken } })
                    .then(function (response) {
                    if (response.ok) {
                        return response.json();
                    }
                    else {
                        _this.setState({
                            error: response.status + " " + response.statusText
                        });
                        throw Error(response.status + " " + response.statusText);
                    }
                })
                    .then(function (response) {
                    _this.setState({
                        items: response.filter(function (i) { return i.start && (new Date(i.start) >= new Date(new Date().setHours(-24))); }),
                        initialized: true
                    });
                });
        }).catch(function (err) {
            _this.setState({
                error: err.message
            });
        });
    };
    return JobStatistics;
}(React.Component));
exports.JobStatistics = JobStatistics;
//# sourceMappingURL=JobStatistics.js.map