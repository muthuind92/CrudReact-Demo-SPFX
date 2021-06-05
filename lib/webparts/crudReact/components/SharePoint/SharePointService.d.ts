import { WebPartContext } from "@microsoft/sp-webpart-base";
import { EnvironmentType } from "@microsoft/sp-core-library";
import { IListCollection } from "./IList";
import { IListItemCollection } from "./IListItem";
import { IListFieldCollection } from "./IListField";
export declare class SharePointServiceManager {
    context: WebPartContext;
    environmentType: EnvironmentType;
    setup(context: WebPartContext, environmentType: EnvironmentType): void;
    get(relativeEndpointUrl: string): Promise<any>;
    getLists(showHiddenLists?: boolean): Promise<IListCollection>;
    getListItems(listId: string, selectedFields?: string[]): Promise<IListItemCollection>;
    getListFields(listId: string, showHiddenFields?: boolean): Promise<IListFieldCollection>;
}
declare const SharePointService: SharePointServiceManager;
export default SharePointService;
//# sourceMappingURL=SharePointService.d.ts.map