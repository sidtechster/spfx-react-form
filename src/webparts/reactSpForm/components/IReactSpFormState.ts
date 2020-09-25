import { IListItem } from './IListItem';

export interface IReactSpFormState {
    status: string;
    listItems: IListItem[],
    listItem: IListItem
}