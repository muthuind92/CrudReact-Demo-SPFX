var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import * as React from 'react';
import styles from './CrudReact.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient } from '@microsoft/sp-http';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Dropdown, DropdownMenuItemType } from 'office-ui-fabric-react/lib/Dropdown';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { PeoplePicker } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { sp } from "@pnp/sp";
import { Stack } from 'office-ui-fabric-react/lib/Stack';
import pnp from "sp-pnp-js";
import { PrimaryButton } from '@fluentui/react/lib/Button';
import { DefaultButton } from '@fluentui/react/lib/Button';
var stackTokens = { childrenGap: 20 };
var dropdownStyles = {
    dropdown: { width: 300 },
};
var drpitems = [];
var CrudReact = /** @class */ (function (_super) {
    __extends(CrudReact, _super);
    function CrudReact(props, state) {
        var _this = _super.call(this, props) || this;
        _this.handleTitle = _this.handleTitle.bind(_this);
        _this.handleDesc = _this.handleDesc.bind(_this);
        _this.AssignedTo = _this.AssignedTo.bind(_this);
        sp.setup({
            spfxContext: _this.props.context
        });
        _this.state = {
            status: 'Ready',
            items: [],
            name: "",
            description: "",
            required: "This is required",
            onSubmission: false,
            AssignedTo: "",
            disableToggle: false,
            defaultChecked: false,
            users: [],
            userManagerIDs: [],
            drpitems: [],
            termnCond: false,
        };
        return _this;
    }
    CrudReact.prototype.render = function () {
        var _this = this;
        var items = this.state.items.map(function (item, i) {
            return (React.createElement("li", null,
                item.Title,
                " (",
                item.Id,
                ") "));
        });
        return (React.createElement("div", { className: styles.crudReact },
            React.createElement("div", { className: styles.container },
                React.createElement("div", { className: styles.row },
                    React.createElement("div", { className: styles.column },
                        React.createElement("p", { className: styles.description }, escape(this.props.listName)),
                        React.createElement("div", { className: "ms-Grid-row ms-fontColor-white " + styles.row },
                            React.createElement("div", { className: 'ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1' },
                                React.createElement("a", { href: "#", className: "" + styles.button, onClick: function () { return _this.createItem(); } },
                                    React.createElement("span", { className: styles.label }, "Create an item")),
                                React.createElement("a", { href: "#", className: "" + styles.button, onClick: function () { return _this.readItem(); } },
                                    React.createElement("span", { className: styles.label }, "Read an item")))),
                        React.createElement("div", { className: "ms-Grid-row ms-bgColor-themeDark ms-fontColor-white " + styles.row },
                            React.createElement("div", { className: 'ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1' },
                                React.createElement("a", { href: "#", className: "" + styles.button, onClick: function () { return _this.updateItem(); } },
                                    React.createElement("span", { className: styles.label }, "Update an item")),
                                React.createElement("a", { href: "#", className: "" + styles.button, onClick: function () { return _this.deleteItem(); } },
                                    React.createElement("span", { className: styles.label }, "Delete an item")))),
                        React.createElement("div", { className: "ms-Grid-row ms-bgColor-themeDark ms-fontColor-white " + styles.row },
                            React.createElement("div", { className: 'ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1' },
                                this.state.status,
                                React.createElement("ul", null, items))))),
                React.createElement("div", { className: styles.row },
                    React.createElement("div", { className: styles.column },
                        React.createElement("p", { className: styles.description }, "Registration Form"),
                        React.createElement("div", { className: "ms-Grid-col ms-u-sm4 block" },
                            React.createElement("label", { className: "ms-Label" }, "Employee Name")),
                        React.createElement("div", { className: "ms-Grid-col ms-u-sm8 block" },
                            React.createElement(TextField, { value: this.state.name, required: true, onChanged: this.handleTitle, errorMessage: (this.state.name.length === 0 && this.state.onSubmission === true) ? this.state.required : "" })),
                        React.createElement("div", { className: "ms-Grid-col ms-u-sm4 block" },
                            React.createElement("label", { className: "ms-Label" }, "Job Description")),
                        React.createElement("div", { className: "ms-Grid-col ms-u-sm8 block" },
                            React.createElement(TextField, { multiline: true, autoAdjustHeight: true, value: this.state.description, onChanged: this.handleDesc })),
                        React.createElement("div", { className: "ms-Grid-col ms-u-sm4 block" },
                            React.createElement("label", { className: "ms-Label" }, "Project Assigned To"),
                            React.createElement("br", null)),
                        React.createElement("div", { className: "ms-Grid-col ms-u-sm8 block" },
                            React.createElement(TextField, { value: this.state.AssignedTo, required: true, onChanged: this.AssignedTo, errorMessage: (this.state.name.length === 0 && this.state.onSubmission === true) ? this.state.required : "" })),
                        React.createElement("div", { className: "ms-Grid-col ms-u-sm4 block" },
                            React.createElement("label", { className: "ms-Label" }, "External Hiring?")),
                        React.createElement("div", { className: "ms-Grid-col ms-u-sm8 block" },
                            React.createElement(Toggle, { disabled: this.state.disableToggle, checked: this.state.defaultChecked, label: "", onAriaLabel: "This toggle is checked. Press to uncheck.", offAriaLabel: "This toggle is unchecked. Press to check.", onText: "On", offText: "Off", onChanged: function (checked) { return _this._changeSharing(checked); }, onFocus: function () { return console.log('onFocus called'); }, onBlur: function () { return console.log('onBlur called'); } })),
                        React.createElement("div", { className: "ms-Grid-col ms-u-sm4 block" },
                            React.createElement("label", { className: "ms-Label" }, "Reporting Manager")),
                        React.createElement("div", null,
                            React.createElement(PeoplePicker, { context: this.props.context, titleText: " ", personSelectionLimit: 1, groupName: "", showtooltip: false, required: true, disabled: false, errorMessage: (this.state.userManagerIDs.length === 0 && this.state.onSubmission === true) ? this.state.required : " " })),
                        React.createElement("div", { className: "ms-Grid-col ms-u-sm4 block" },
                            React.createElement("label", { className: "ms-Label" }, "Department"),
                            React.createElement("br", null)),
                        React.createElement("div", { className: "ms-Grid-col ms-u-sm8 block" },
                            React.createElement(Stack, { tokens: stackTokens },
                                React.createElement(Dropdown, { placeholder: "Select an option", options: this.state.drpitems, styles: dropdownStyles }))),
                        React.createElement("div", { className: "ms-Grid-col ms-u-sm6 block" }),
                        React.createElement("div", { className: "ms-Grid-col ms-u-sm2 block" },
                            React.createElement(PrimaryButton, { text: "Create", onClick: function () { _this.validateForm(); } })),
                        React.createElement("div", { className: "ms-Grid-col ms-u-sm2 block" },
                            React.createElement(DefaultButton, { text: "Cancel", onClick: function () { _this.setState({}); } })))))));
    };
    CrudReact.prototype.getLatestItemId = function () {
        var _this = this;
        return new Promise(function (resolve, reject) {
            _this.props.spHttpClient.get(_this.props.siteUrl + "/_api/web/lists/getbytitle('" + _this.props.listName + "')/items?$orderby=Id desc&$top=1&$select=id", SPHttpClient.configurations.v1, {
                headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'odata-version': ''
                }
            })
                .then(function (response) {
                return response.json();
            }, function (error) {
                reject(error);
            })
                .then(function (response) {
                if (response.value.length === 0) {
                    resolve(-1);
                }
                else {
                    resolve(response.value[0].Id);
                }
            });
        });
    };
    CrudReact.prototype.createItem = function () {
        var _this = this;
        this.setState({
            status: 'Creating item...',
            items: []
        });
        var body = JSON.stringify({
            'Title': "Test item created by SPFx ReactJS on: " + new Date()
        });
        this.props.spHttpClient.post(this.props.siteUrl + "/_api/web/lists/getbytitle('" + this.props.listName + "')/items", SPHttpClient.configurations.v1, {
            headers: {
                'Accept': 'application/json;odata=nometadata',
                'Content-type': 'application/json;odata=nometadata',
                'odata-version': ''
            },
            body: body
        })
            .then(function (response) {
            return response.json();
        })
            .then(function (item) {
            _this.setState({
                status: "Item '" + item.Title + "' (ID: " + item.Id + ") successfully created",
                items: []
            });
        }, function (error) {
            _this.setState({
                status: 'Error while creating the item: ' + error,
                items: []
            });
        });
    };
    CrudReact.prototype.readItem = function () {
        var _this = this;
        this.setState({
            status: 'Loading latest items...',
            items: []
        });
        this.getLatestItemId()
            .then(function (itemId) {
            if (itemId === -1) {
                throw new Error('No items found in the list');
            }
            _this.setState({
                status: "Loading information about item ID: " + itemId + "...",
                items: []
            });
            return _this.props.spHttpClient.get(_this.props.siteUrl + "/_api/web/lists/getbytitle('" + _this.props.listName + "')/items(" + itemId + ")?$select=Title,Id", SPHttpClient.configurations.v1, {
                headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'odata-version': ''
                }
            });
        })
            .then(function (response) {
            return response.json();
        })
            .then(function (item) {
            _this.setState({
                status: "Item ID: " + item.Id + ", Title: " + item.Title,
                items: []
            });
        }, function (error) {
            _this.setState({
                status: 'Loading latest item failed with error: ' + error,
                items: []
            });
        });
    };
    //#region 
    CrudReact.prototype.updateItem = function () {
        var _this = this;
        this.setState({
            status: 'Loading latest items...',
            items: []
        });
        var latestItemId = undefined;
        this.getLatestItemId()
            .then(function (itemId) {
            if (itemId === -1) {
                throw new Error('No items found in the list');
            }
            latestItemId = itemId;
            _this.setState({
                status: "Loading information about item ID: " + latestItemId + "...",
                items: []
            });
            return _this.props.spHttpClient.get(_this.props.siteUrl + "/_api/web/lists/getbytitle('" + _this.props.listName + "')/items(" + latestItemId + ")?$select=Title,Id", SPHttpClient.configurations.v1, {
                headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'odata-version': ''
                }
            });
        })
            .then(function (response) {
            return response.json();
        })
            .then(function (item) {
            _this.setState({
                status: 'Loading latest items...',
                items: []
            });
            var body = JSON.stringify({
                'Title': "Updated Item " + new Date()
            });
            _this.props.spHttpClient.post(_this.props.siteUrl + "/_api/web/lists/getbytitle('" + _this.props.listName + "')/items(" + item.Id + ")", SPHttpClient.configurations.v1, {
                headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'Content-type': 'application/json;odata=nometadata',
                    'odata-version': '',
                    'IF-MATCH': '*',
                    'X-HTTP-Method': 'MERGE'
                },
                body: body
            })
                .then(function (response) {
                _this.setState({
                    status: "Item with ID: " + latestItemId + " successfully updated",
                    items: []
                });
            }, function (error) {
                _this.setState({
                    status: "Error updating item: " + error,
                    items: []
                });
            });
        });
    };
    //#endregion
    CrudReact.prototype.deleteItem = function () {
        var _this = this;
        if (!window.confirm('Are you sure you want to delete the latest item?')) {
            return;
        }
        this.setState({
            status: 'Loading latest items...',
            items: []
        });
        var latestItemId = undefined;
        var etag = undefined;
        this.getLatestItemId()
            .then(function (itemId) {
            if (itemId === -1) {
                throw new Error('No items found in the list');
            }
            latestItemId = itemId;
            _this.setState({
                status: "Loading information about item ID: " + latestItemId + "...",
                items: []
            });
            return _this.props.spHttpClient.get(_this.props.siteUrl + "/_api/web/lists/getbytitle('" + _this.props.listName + "')/items(" + latestItemId + ")?$select=Id", SPHttpClient.configurations.v1, {
                headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'odata-version': ''
                }
            });
        })
            .then(function (response) {
            etag = response.headers.get('ETag');
            return response.json();
        })
            .then(function (item) {
            _this.setState({
                status: "Deleting item with ID: " + latestItemId + "...",
                items: []
            });
            return _this.props.spHttpClient.post(_this.props.siteUrl + "/_api/web/lists/getbytitle('" + _this.props.listName + "')/items(" + item.Id + ")", SPHttpClient.configurations.v1, {
                headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'Content-type': 'application/json;odata=verbose',
                    'odata-version': '',
                    'IF-MATCH': etag,
                    'X-HTTP-Method': 'DELETE'
                }
            });
        })
            .then(function (response) {
            _this.setState({
                status: "Item with ID: " + latestItemId + " successfully deleted",
                items: []
            });
        }, function (error) {
            _this.setState({
                status: "Error deleting item: " + error,
                items: []
            });
        });
    };
    //#region  Registration Form Methods 
    CrudReact.prototype.handleTitle = function (value) {
        return this.setState({
            name: value
        });
    };
    CrudReact.prototype.handleDesc = function (value) {
        return this.setState({
            description: value
        });
    };
    CrudReact.prototype.AssignedTo = function (value) {
        return this.setState({
            AssignedTo: value
        });
    };
    CrudReact.prototype._changeSharing = function (checked) {
        this.setState({ defaultChecked: checked });
    };
    CrudReact.prototype._getPeoplePickerItems = function (items) {
        console.log('Items:', items);
        this.setState({ users: items });
    };
    CrudReact.prototype._log = function (str) {
        return function () {
            console.log(str);
        };
    };
    CrudReact.prototype._onCheckboxChange = function (ev, isChecked) {
        console.log("The option has been changed to " + isChecked + ".");
        //this.setState({termnCond: (isChecked)?true:false});
    };
    CrudReact.prototype.validateForm = function () {
        var allowCreate = true;
        this.setState({ onSubmission: true });
        if (this.state.name.length === 0) {
            allowCreate = false;
        }
        //if(this.state.termKey === undefined)
        //{
        //  allowCreate = false;
        //}   
        if (allowCreate) {
            //this._onShowPanel();
        }
        else {
            //do nothing
        }
    };
    CrudReact.prototype.componentDidMount = function () {
        return __awaiter(this, void 0, void 0, function () {
            var reacthandler;
            return __generator(this, function (_a) {
                reacthandler = this;
                pnp.sp.web.lists
                    .getByTitle("Department_Master")
                    .items.select("Title")
                    .get()
                    .then(function (data) {
                    drpitems.push({ key: 'Department Header', text: 'Department', itemType: DropdownMenuItemType.Header });
                    for (var k in data) {
                        drpitems.push({ key: data[k].Title, text: data[k].Title });
                    }
                    reacthandler.setState({ drpitems: drpitems });
                    console.log(drpitems);
                    return drpitems;
                });
                return [2 /*return*/];
            });
        });
    };
    return CrudReact;
}(React.Component));
export default CrudReact;
//# sourceMappingURL=CrudReact.js.map