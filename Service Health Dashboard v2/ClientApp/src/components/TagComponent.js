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
exports.Tag = exports.TagViewState = void 0;
var React = require("react");
var react_1 = require("@fluentui/react");
var GlobalState_1 = require("./GlobalState");
var iconClass = react_1.mergeStyles({
    fontSize: 16,
    height: 16,
    width: 16,
    margin: '-4px 0px 4px 4px'
});
var classNames = react_1.mergeStyleSets({
    New: [{ color: '#0078d4' }, iconClass],
    Retirement: [{ color: '#d83b01' }, iconClass],
    Updated: [{ color: '#31752f' }, iconClass],
    Warning: [{ color: '#f2c80f' }, iconClass],
    Admin: [{ color: '#007fff' }, iconClass],
    AdminImpact: [{ color: '#d83b01' }, iconClass],
    Default: [{ color: 'black' }, iconClass],
});
var serviceMapping = [
    { name: 'New Feature', icon: 'AddIn', color: classNames.New },
    { name: 'Retirement', icon: 'PageRemove', color: classNames.Retirement },
    { name: 'Updated message', icon: 'Message', color: classNames.Updated },
    { name: 'User impact', icon: 'UserWarning', color: classNames.Warning },
    { name: 'Feature update', icon: 'ProductRelease', color: classNames.Updated },
    { name: 'Admin impact', icon: 'Admin', color: classNames.AdminImpact }
];
var TagViewState;
(function (TagViewState) {
    TagViewState[TagViewState["ShowAll"] = 0] = "ShowAll";
    TagViewState[TagViewState["IconOnly"] = 1] = "IconOnly";
    TagViewState[TagViewState["TextOnly"] = 2] = "TextOnly";
})(TagViewState = exports.TagViewState || (exports.TagViewState = {}));
var Tag = /** @class */ (function (_super) {
    __extends(Tag, _super);
    function Tag(props) {
        var _this = _super.call(this, props) || this;
        _this.state = {
            name: "",
            compact: true,
            viewState: TagViewState.ShowAll,
            iconTitle: undefined,
            tagMapping: { name: "", icon: "", color: classNames.Default }
        };
        return _this;
    }
    Tag.prototype.render = function () {
        var _a = this.state, name = _a.name, compact = _a.compact, tagMapping = _a.tagMapping, viewState = _a.viewState, iconTitle = _a.iconTitle;
        var globalState = this.context;
        var theme = globalState.getTheme();
        var margins = compact === true ? '0 4px' : '4px 4px 4px';
        var border = viewState === TagViewState.IconOnly ? "" : "1px solid " + theme.semanticColors.bodyDivider;
        var background = viewState === TagViewState.IconOnly ? "" : theme.semanticColors.menuBackground;
        if (name === undefined)
            return (React.createElement("div", null));
        else
            return (React.createElement("div", { className: 'tagComponent', style: { border: border, alignItems: 'center', display: 'inline-flex', margin: margins, background: background } },
                React.createElement(react_1.FontIcon, { iconName: tagMapping.icon, className: tagMapping.color, hidden: tagMapping.icon === '' || !(viewState === TagViewState.ShowAll || viewState === TagViewState.IconOnly), title: iconTitle }),
                React.createElement(react_1.Text, { variant: 'xSmall', style: { margin: '-4px 4px 4px 4px', paddingTop: '8px' }, hidden: name === null ? true : name.toLowerCase().trim() === '' || !(viewState === TagViewState.ShowAll || viewState === TagViewState.TextOnly) }, name === null ? "" : (name.toLowerCase().trim() === '' || !(viewState === TagViewState.ShowAll || viewState === TagViewState.TextOnly) ? "" : name.toUpperCase()))));
    };
    Tag.prototype.componentDidMount = function () {
        this._setComponentState();
    };
    Tag.prototype.componentDidUpdate = function (prevProps) {
        if (this.props.name != prevProps.name ||
            this.props.compact != prevProps.compact ||
            this.props.viewState != prevProps.viewState ||
            this.props.iconTitle != prevProps.iconTitle) {
            this._setComponentState();
        }
    };
    Tag.prototype._setComponentState = function () {
        var _this = this;
        var mappingMatch = this.props.name === undefined || this.props.name === null ? undefined : serviceMapping.find(function (m) { return m.name.trim().toLowerCase() === _this.props.name.trim().toLowerCase(); });
        var defaultMapping = { name: '', icon: '', color: classNames.Default };
        this.setState({
            name: this.props.name === null ? "" : this.props.name,
            compact: this.props.compact !== undefined ? this.props.compact : true,
            viewState: this.props.viewState !== undefined ? this.props.viewState : TagViewState.ShowAll,
            iconTitle: this.props.iconTitle,
            tagMapping: mappingMatch ? mappingMatch : defaultMapping
        });
    };
    Tag.contextType = GlobalState_1.GlobalState;
    return Tag;
}(React.Component));
exports.Tag = Tag;
//# sourceMappingURL=TagComponent.js.map