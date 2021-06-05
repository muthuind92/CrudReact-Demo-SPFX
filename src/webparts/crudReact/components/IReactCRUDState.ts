import { IListItem } from './IListItem';  
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';

export interface IReactCRUDState {  
  status: string;  
  items: IListItem[];  
  name:string;
  description:string;
  onSubmission:boolean;
  required:string;
  AssignedTo:string;
  disableToggle:boolean,
  defaultChecked:boolean,
  users: any[]; 
  userManagerIDs: number[];
  drpitems: IDropdownOption[] ,
  termnCond:boolean,
  
}  