import { SPHttpClient } from '@microsoft/sp-http'; 
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ICrudReactProps
 {
   
   listName: string;   
   spHttpClient: SPHttpClient;  
   siteUrl: string; 
   description: string;
   context: WebPartContext;
   
}
