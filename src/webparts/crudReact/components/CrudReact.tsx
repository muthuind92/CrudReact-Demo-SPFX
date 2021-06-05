import * as React from 'react';
import styles from './CrudReact.module.scss';
import { ICrudReactProps } from './ICrudReactProps';
import { IReactCRUDState } from './IReactCRUDState';
import { escape } from '@microsoft/sp-lodash-subset';
import { IListItem } from './IListItem';  
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';  
import { TextField } from 'office-ui-fabric-react/lib/TextField'; 
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { sp } from "@pnp/sp";
import { IStackTokens, Stack } from 'office-ui-fabric-react/lib/Stack';
import pnp from "sp-pnp-js";
import SharePointService from './SharePoint/SharePointService';
import {Checkbox} from 'office-ui-fabric-react/lib/Checkbox'
import { PrimaryButton } from '@fluentui/react/lib/Button';
import {DefaultButton} from '@fluentui/react/lib/Button';

const stackTokens: IStackTokens = { childrenGap: 20 };

const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: 300 },
};


var drpitems: IDropdownOption[] = [] ;


export default class CrudReact extends React.Component<ICrudReactProps, IReactCRUDState> {


 
  constructor(props: ICrudReactProps, state: IReactCRUDState) {  
    super(props);  
    this.handleTitle = this.handleTitle.bind(this);
    this.handleDesc = this.handleDesc.bind(this);
    this.AssignedTo=this.AssignedTo.bind(this);
    
    sp.setup({
      spfxContext: this.props.context
    });
    this.state = {  
      status: 'Ready',  
      items: [] ,
      name:"",
      description:"",
      required:"This is required",
      onSubmission:false,
      AssignedTo:"",
      disableToggle:false,
      defaultChecked:false,
      users: []  ,
      userManagerIDs: [],
      drpitems:[],
      termnCond:false,
     
      
    };  
  }  
  

  public render(): React.ReactElement<ICrudReactProps>{  
  
  

    const items: JSX.Element[] = this.state.items.map((item: IListItem, i: number): JSX.Element => {  
      return (  
        <li>{item.Title} ({item.Id}) </li>  
      );  
    });  
  
    
    return ( 
      
      
      <div className={ styles.crudReact }>  
        <div className={ styles.container }>  
          <div className={ styles.row }>  
        
            <div className={ styles.column }>  
              
           <p className={ styles.description }>{escape(this.props.listName)}</p>  
           

              <div className={`ms-Grid-row ms-fontColor-white ${styles.row}`}>  
                <div className='ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1'>  
                  <a href="#" className={`${styles.button}`} onClick={() => this.createItem()}>  
                    <span className={styles.label}>Create an item</span>  
                  </a>   
                  <a href="#" className={`${styles.button}`} onClick={() => this.readItem()}>  
                    <span className={styles.label}>Read an item</span>  
                  </a>  
                </div>  
              </div>  
  
              <div className={`ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}`}>  
                <div className='ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1'>  
                  <a href="#" className={`${styles.button}`} onClick={() => this.updateItem()}>  
                    <span className={styles.label}>Update an item</span>  
                  </a>   
                  <a href="#" className={`${styles.button}`} onClick={() => this.deleteItem()}>  
                    <span className={styles.label}>Delete an item</span>  
                  </a>  
                </div>  
              </div>  
  
              <div className={`ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}`}>  
                <div className='ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1'>  
                  {this.state.status}  
                  <ul>  
                    {items}  
                  </ul>  
                </div>  
              </div>  
  
            </div>  
          </div>  

          <div className={ styles.row }>  
            <div className={ styles.column }> 
            <p className={ styles.description }>Registration Form</p> 
            <div className="ms-Grid-col ms-u-sm4 block">
              <label className="ms-Label">Employee Name</label>             
         </div>
         <div className="ms-Grid-col ms-u-sm8 block">
             <TextField value={this.state.name} required={true} onChanged={this.handleTitle}
         errorMessage={(this.state.name.length === 0 && this.state.onSubmission === true) ? this.state.required : ""}/>
        </div>
        <div className="ms-Grid-col ms-u-sm4 block">
             <label className="ms-Label">Job Description</label>
          </div>
          <div className="ms-Grid-col ms-u-sm8 block">
             <TextField multiline autoAdjustHeight value={this.state.description} onChanged={this.handleDesc}
              />
          </div>
          <div className="ms-Grid-col ms-u-sm4 block">
             <label className="ms-Label">Project Assigned To</label><br/>
              
          </div>
          <div className="ms-Grid-col ms-u-sm8 block">
             <TextField value={this.state.AssignedTo} required={true} onChanged={this.AssignedTo}
         errorMessage={(this.state.name.length === 0 && this.state.onSubmission === true) ? this.state.required : ""}/>
        </div>
         
        <div className="ms-Grid-col ms-u-sm4 block">
          <label className="ms-Label">External Hiring?</label>
        </div>
        <div className="ms-Grid-col ms-u-sm8 block">
        <Toggle
          disabled={this.state.disableToggle}
          checked={this.state.defaultChecked}
          label=""
          onAriaLabel="This toggle is checked. Press to uncheck."
          offAriaLabel="This toggle is unchecked. Press to check."
          onText="On"
          offText="Off"
          onChanged={(checked) =>this._changeSharing(checked)}
          onFocus={() => console.log('onFocus called')}
          onBlur={() => console.log('onBlur called')}         
        />
        </div>
        
        <div className="ms-Grid-col ms-u-sm4 block">
          <label className="ms-Label">Reporting Manager</label>
        </div>
        <div>
        <PeoplePicker
            context={this.props.context as any}
            titleText=" "
            personSelectionLimit={1}
            groupName={""} // Leave this blank in casessss you want to filter from all users
            showtooltip={false}
            required={true}
            disabled={false}
            errorMessage={(this.state.userManagerIDs.length === 0 && this.state.onSubmission === true) ? this.state.required : " "}
            
            />
         </div>

         <div className="ms-Grid-col ms-u-sm4 block">
             <label className="ms-Label">Department</label><br/>
          </div>
          <div className="ms-Grid-col ms-u-sm8 block">
          <Stack tokens={stackTokens}>
          <Dropdown
          placeholder="Select an option"
          options={this.state.drpitems}
          
          styles={dropdownStyles}
         
        />
              </Stack>

        </div>
        <div className="ms-Grid-col ms-u-sm6 block">
         </div>
         <div className="ms-Grid-col ms-u-sm2 block">
           <PrimaryButton text="Create" onClick={() => { this.validateForm(); }} />
        </div>
        <div className="ms-Grid-col ms-u-sm2 block">
           <DefaultButton text="Cancel" onClick={() => { this.setState({}); }} />
        </div>
    
           
         

</div>
</div>
        </div>
        </div>  
    
    
    
    
    );  
  }  


  private getLatestItemId(): Promise<number> {  
    return new Promise<number>((resolve: (itemId: number) => void, reject: (error: any) => void): void => {  
      this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items?$orderby=Id desc&$top=1&$select=id`,  
        SPHttpClient.configurations.v1,  
        {  
          headers: {  
            'Accept': 'application/json;odata=nometadata',  
            'odata-version': ''  
          }  
        })  
        .then((response: SPHttpClientResponse): Promise<{ value: { Id: number }[] }> => {  
          return response.json();  
        }, (error: any): void => {  
          reject(error);  
        })  
        .then((response: { value: { Id: number }[] }): void => {  
          if (response.value.length === 0) {  
            resolve(-1);  
          }  
          else {  
            resolve(response.value[0].Id);  
          }  
        });  
    });  
  }   


  private createItem(): void {  
    this.setState({  
      status: 'Creating item...',  
      items: []  
    });  
    
    const body: string = JSON.stringify({  
      'Title': `Test item created by SPFx ReactJS on: ${new Date()}`  
    });  
    
    this.props.spHttpClient.post(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items`,  
    SPHttpClient.configurations.v1,  
    {  
      headers: {  
        'Accept': 'application/json;odata=nometadata',  
        'Content-type': 'application/json;odata=nometadata',  
        'odata-version': ''  
      },  
      body: body  
    })  
    .then((response: SPHttpClientResponse): Promise<IListItem> => {  
      return response.json();  
    })  
    .then((item: IListItem): void => {  
      this.setState({  
        status: `Item '${item.Title}' (ID: ${item.Id}) successfully created`,  
        items: []  
      });  
    }, (error: any): void => {  
      this.setState({  
        status: 'Error while creating the item: ' + error,  
        items: []  
      });  
    });  
  }  
  
  private readItem(): void {  
    this.setState({  
      status: 'Loading latest items...',  
      items: []  
    });  
    
    this.getLatestItemId()  
      .then((itemId: number): Promise<SPHttpClientResponse> => {  
        if (itemId === -1) {  
          throw new Error('No items found in the list');  
        }  
    
        this.setState({  
          status: `Loading information about item ID: ${itemId}...`,  
          items: []  
        });  
        return this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items(${itemId})?$select=Title,Id`,  
          SPHttpClient.configurations.v1,  
          {  
            headers: {  
              'Accept': 'application/json;odata=nometadata',  
              'odata-version': ''  
            }  
          });  
      })  
      .then((response: SPHttpClientResponse): Promise<IListItem> => {  
        return response.json();  
      })  
      .then((item: IListItem): void => {  
        this.setState({  
          status: `Item ID: ${item.Id}, Title: ${item.Title}`,  
          items: []  
        });  
      }, (error: any): void => {  
        this.setState({  
          status: 'Loading latest item failed with error: ' + error,  
          items: []  
        });  
      });  
  }  
  //#region 
  private updateItem(): void {  
    this.setState({  
      status: 'Loading latest items...',  
      items: []  
    });  
    
    let latestItemId: number = undefined;  
    
    this.getLatestItemId()  
      .then((itemId: number): Promise<SPHttpClientResponse> => {  
        if (itemId === -1) {  
          throw new Error('No items found in the list');  
        }  
    
        latestItemId = itemId;  
        this.setState({  
          status: `Loading information about item ID: ${latestItemId}...`,  
          items: []  
        });  
          
        return this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items(${latestItemId})?$select=Title,Id`,  
          SPHttpClient.configurations.v1,  
          {  
            headers: {  
              'Accept': 'application/json;odata=nometadata',  
              'odata-version': ''  
            }  
          });  
      })  
      .then((response: SPHttpClientResponse): Promise<IListItem> => {  
        return response.json();  
      })  
      .then((item: IListItem): void => {  
        this.setState({  
          status: 'Loading latest items...',  
          items: []  
        });  
    
        const body: string = JSON.stringify({  
          'Title': `Updated Item ${new Date()}`  
        });  
    
        this.props.spHttpClient.post(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items(${item.Id})`,  
          SPHttpClient.configurations.v1,  
          {  
            headers: {  
              'Accept': 'application/json;odata=nometadata',  
              'Content-type': 'application/json;odata=nometadata',  
              'odata-version': '',  
              'IF-MATCH': '*',  
              'X-HTTP-Method': 'MERGE'  
            },  
            body: body  
          })  
          .then((response: SPHttpClientResponse): void => {  
            this.setState({  
              status: `Item with ID: ${latestItemId} successfully updated`,  
              items: []  
            });  
          }, (error: any): void => {  
            this.setState({  
              status: `Error updating item: ${error}`,  
              items: []  
            });  
          });  
      });  
  }   
  //#endregion
  
  


  private deleteItem(): void
   {  
    if (!window.confirm('Are you sure you want to delete the latest item?'))
     {  
      return;  
     }  
    
    this.setState({  
      status: 'Loading latest items...',  
      items: []  
    });  
    
    let latestItemId: number = undefined;  
    let etag: string = undefined;  
    this.getLatestItemId()  
      .then((itemId: number): Promise<SPHttpClientResponse> => {  
        if (itemId === -1) {  
          throw new Error('No items found in the list');  
        }  
    
        latestItemId = itemId;  
        this.setState({  
          status: `Loading information about item ID: ${latestItemId}...`,  
          items: []  
        });  
    
        return this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items(${latestItemId})?$select=Id`,  
          SPHttpClient.configurations.v1,  
          {  
            headers: {  
              'Accept': 'application/json;odata=nometadata',  
              'odata-version': ''  
            }  
          });  
      })  
      .then((response: SPHttpClientResponse): Promise<IListItem> => {  
        etag = response.headers.get('ETag');  
        return response.json();  
      })  
      .then((item: IListItem): Promise<SPHttpClientResponse> => {  
        this.setState({  
          status: `Deleting item with ID: ${latestItemId}...`,  
          items: []  
        });  
    
        return this.props.spHttpClient.post(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items(${item.Id})`,  
          SPHttpClient.configurations.v1,  
          {  
            headers: {  
              'Accept': 'application/json;odata=nometadata',  
              'Content-type': 'application/json;odata=verbose',  
              'odata-version': '',  
              'IF-MATCH': etag,  
              'X-HTTP-Method': 'DELETE'  
            }  
          });  
      })  
      .then((response: SPHttpClientResponse): void => {  
        this.setState({  
          status: `Item with ID: ${latestItemId} successfully deleted`,  
          items: []  
        });  
      }, (error: any): void => {  
        this.setState({  
          status: `Error deleting item: ${error}`,  
          items: []  
        });  
      });  
  }  
  

    //#region  Registration Form Methods 
  private handleTitle(value: string): void {
    return this.setState({
      name: value
    });
  }

  private handleDesc(value: string): void {
    return this.setState({
      description: value
    });
  }
  private AssignedTo(value: string): void {
    return this.setState({
      AssignedTo: value
    });
  }
  
  private _changeSharing(checked:any):void{
    this.setState({defaultChecked: checked});
  }
  
  private _getPeoplePickerItems(items: any[]) {  
    console.log('Items:', items);  
    this.setState({ users: items });
  }
  
  
  private _log(str: string): () => void {
    return (): void => {
      console.log(str);
    };
  }
  private _onCheckboxChange(ev: React.FormEvent<HTMLElement>, isChecked: boolean): void {
    console.log(`The option has been changed to ${isChecked}.`);
    //this.setState({termnCond: (isChecked)?true:false});
  }
  private validateForm():void{
    let allowCreate: boolean = true;
    this.setState({ onSubmission : true });
     
    if(this.state.name.length === 0)
    {
      allowCreate = false;
    }
    //if(this.state.termKey === undefined)
    //{
    //  allowCreate = false;
    //}   
    
    if(allowCreate)
    {
       //this._onShowPanel();
    }
    else
    {
      //do nothing
    } 
  }
 
  
  
  public async componentDidMount(): Promise<void> {
    // get all the items from a sharepoint list
    var reacthandler = this;
    pnp.sp.web.lists
      .getByTitle("Department_Master")
      .items.select("Title")
      .get()
      .then(function (data) {
        drpitems.push({key: 'Department Header', text: 'Department', itemType: DropdownMenuItemType.Header})
        for (var k in data) {
         
          drpitems.push({ key: data[k].Title, text: data[k].Title });
        }
        
        reacthandler.setState({ drpitems});
        console.log(drpitems);
        return drpitems;
      });
  }
  
  //#endregion


}
