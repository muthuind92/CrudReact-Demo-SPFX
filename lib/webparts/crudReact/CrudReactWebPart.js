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
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Environment } from '@microsoft/sp-core-library';
import { PropertyPaneDropdown } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'CrudReactWebPartStrings';
import CrudReact from './components/CrudReact';
import SharePointService from './components/SharePoint/SharePointService';
var CrudReactWebPart = /** @class */ (function (_super) {
    __extends(CrudReactWebPart, _super);
    function CrudReactWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.listOptionsLoading = false;
        return _this;
    }
    CrudReactWebPart.prototype.render = function () {
        var element = React.createElement(CrudReact, {
            listName: this.properties.listName,
            spHttpClient: this.context.spHttpClient,
            siteUrl: this.context.pageContext.web.absoluteUrl,
            description: this.properties.description,
            context: this.context
        });
        ReactDom.render(element, this.domElement);
    };
    CrudReactWebPart.prototype.onInit = function () {
        var _this = this;
        return _super.prototype.onInit.call(this).then(function () {
            SharePointService.setup(_this.context, Environment.type);
        });
    };
    CrudReactWebPart.prototype.onDispose = function () {
        ReactDom.unmountComponentAtNode(this.domElement);
    };
    Object.defineProperty(CrudReactWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: true,
        configurable: true
    });
    CrudReactWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupFields: [
                                PropertyPaneDropdown('listName', {
                                    label: strings.ListNameFieldLabel,
                                    options: this.listOptions,
                                    disabled: this.listOptionsLoading,
                                }),
                            ]
                        },
                    ]
                }
            ]
        };
    };
    CrudReactWebPart.prototype.getLists = function () {
        var _this = this;
        this.listOptionsLoading = true;
        this.context.propertyPane.refresh();
        return SharePointService.getLists().then(function (lists) {
            _this.listOptionsLoading = false;
            _this.context.propertyPane.refresh();
            return lists.value.map(function (list) {
                return {
                    key: list.Title,
                    text: list.Title,
                };
            });
        });
    };
    CrudReactWebPart.prototype.onPropertyPaneConfigurationStart = function () {
        var _this = this;
        this.getLists()
            .then(function (listOptions) {
            _this.listOptions = listOptions;
            _this.context.propertyPane.refresh();
            _this.context.statusRenderer.clearLoadingIndicator(_this.domElement);
            _this.render();
        });
    };
    return CrudReactWebPart;
}(BaseClientSideWebPart));
export default CrudReactWebPart;
//# sourceMappingURL=CrudReactWebPart.js.map