import { IDropdownOption } from 'office-ui-fabric-react';
import { IListItem } from './IListItem';

export interface IReactSpFormState {
    status: string;
    listItems: IListItem[],
    listItem: IListItem,
    employeeTypes: IDropdownOption[]    
}