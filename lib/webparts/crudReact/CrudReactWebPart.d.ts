import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
export interface ICrudReactWebPartProps {
    description: string;
    listName: string;
    name: string;
    status: string;
}
export default class CrudReactWebPart extends BaseClientSideWebPart<ICrudReactWebPartProps> {
    private listOptions;
    private listOptionsLoading;
    render(): void;
    onInit(): Promise<void>;
    protected onDispose(): void;
    protected readonly dataVersion: Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
    private getLists;
    protected onPropertyPaneConfigurationStart(): void;
}
//# sourceMappingURL=CrudReactWebPart.d.ts.map