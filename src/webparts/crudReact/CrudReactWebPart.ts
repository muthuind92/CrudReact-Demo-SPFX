import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version,Environment,EnvironmentType } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneTextField,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'CrudReactWebPartStrings';
import CrudReact from './components/CrudReact';
import { ICrudReactProps } from './components/ICrudReactProps';
import SharePointService from './components/SharePoint/SharePointService';

//List Configuration 
export interface ICrudReactWebPartProps 
{
  description: string;
  listName: string;
  name: string;
  status: string;
 
}


export default class CrudReactWebPart extends BaseClientSideWebPart <ICrudReactWebPartProps> {

  // List options state
  private listOptions: IPropertyPaneDropdownOption[];
  private listOptionsLoading: boolean = false;

 
  

  public render(): void {
    const element: React.ReactElement<ICrudReactProps> = React.createElement(
      CrudReact,
      {
        
        listName: this.properties.listName,
        spHttpClient: this.context.spHttpClient,  
        siteUrl: this.context.pageContext.web.absoluteUrl,
        description: this.properties.description,
        context:this.context
      
      }



      
    );

    ReactDom.render(element, this.domElement);
  }

  public onInit(): Promise<void> {
    return super.onInit().then(() => {
      SharePointService.setup(this.context, Environment.type);
      
    });
  }


  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
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
  }

  private getLists(): Promise<IPropertyPaneDropdownOption[]> {
    this.listOptionsLoading = true;
    this.context.propertyPane.refresh();

    return SharePointService.getLists().then(lists => {
      this.listOptionsLoading = false;
      this.context.propertyPane.refresh();

      return lists.value.map(list => {
        return {
          key: list.Title,
          text: list.Title,
        };
      });
    });
  }

  

  protected onPropertyPaneConfigurationStart(): void {
    this.getLists()
    .then((listOptions: IPropertyPaneDropdownOption[]): void => {
      this.listOptions = listOptions;
      this.context.propertyPane.refresh();
      this.context.statusRenderer.clearLoadingIndicator(this.domElement);
      this.render();
    });
  }

 
}

