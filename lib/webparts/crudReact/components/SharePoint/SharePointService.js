import { SPHttpClient } from "@microsoft/sp-http";
var SharePointServiceManager = /** @class */ (function () {
    function SharePointServiceManager() {
    }
    SharePointServiceManager.prototype.setup = function (context, environmentType) {
        this.context = context;
        this.environmentType = environmentType;
    };
    SharePointServiceManager.prototype.get = function (relativeEndpointUrl) {
        return this.context.spHttpClient.get("" + this.context.pageContext.web.absoluteUrl + relativeEndpointUrl, SPHttpClient.configurations.v1).then(function (response) {
            if (!response.ok)
                return Promise.reject('GET Request Failed');
            return response.json();
        }).catch(function (error) {
            return Promise.reject(error);
        });
    };
    SharePointServiceManager.prototype.getLists = function (showHiddenLists) {
        if (showHiddenLists === void 0) { showHiddenLists = false; }
        return this.get("/_api/lists" + (!showHiddenLists ? '?$filter=Hidden eq false' : ''));
    };
    SharePointServiceManager.prototype.getListItems = function (listId, selectedFields) {
        return this.get("/_api/lists/getbyid('" + listId + "')/items" + (selectedFields ? "?$select=" + selectedFields.join(',') : ''));
    };
    SharePointServiceManager.prototype.getListFields = function (listId, showHiddenFields) {
        if (showHiddenFields === void 0) { showHiddenFields = false; }
        return this.get("/_api/lists/getbyid('" + listId + "')/fields" + (!showHiddenFields ? '?$filter=Hidden eq false' : ''));
    };
    return SharePointServiceManager;
}());
export { SharePointServiceManager };
var SharePointService = new SharePointServiceManager();
export default SharePointService;
//# sourceMappingURL=SharePointService.js.map