import * as React from 'react';
import styles from './ReactSpForm.module.scss';
import { IReactSpFormProps } from './IReactSpFormProps';
import { IReactSpFormState } from './IReactSpFormState';
import { escape } from '@microsoft/sp-lodash-subset';
import { IListItem } from './IListItem';
import { ISPHttpClientOptions, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import { 
  TextField,
  autobind,
  PrimaryButton,
  DetailsList,
  DetailsListLayoutMode,
  CheckboxVisibility,
  SelectionMode,
  Dropdown,
  IDropdown,
  IDropdownOption,
  ITextFieldStyles,
  IDropdownStyles,
  DetailsRowCheck,
  Selection,
  DropdownBase,
  DatePicker,
  IDatePickerStyles,
  IDatePickerStyleProps,
  mergeStyleSets
 } from 'office-ui-fabric-react';

 
 // Configure the columns for the DetailsList component

 let _demoListColumns = [
  {
    key: 'ID',
    name: 'ID',
    fieldName: 'ID',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  },
  {
   key: 'Title',
   name: 'Full Name',
   fieldName: 'Title',
   minWidth: 50,
   maxWidth: 100,
   isResizable: true
 },
 {
  key: 'DateofBirth',
  name: 'Date of Birth',
  fieldName: 'DateofBirth',
  minWidth: 50,
  maxWidth: 100,
  isResizable: true
 },
 {
  key: 'EmployeeType',
  name: 'Employee Type',
  fieldName: 'EmployeeType',
  minWidth: 50,
  maxWidth: 100,
  isResizable: true
 }
];

const controlClass = mergeStyleSets({
  control: {
    maxWidth: '300px'
  }
});

export default class ReactSpForm extends React.Component<IReactSpFormProps, IReactSpFormState> {

  private _selection: Selection;

  private onItemsSelectionChanged = () => {
    this.setState({
      listItem: (this._selection.getSelection()[0] as IListItem)
    });
  }

  constructor(props: IReactSpFormProps, state: IReactSpFormState) {

    super(props);

    let today = new Date();    

    this.state = {
      status: 'Ready',
      listItems: [],
      listItem: {
        Id: 0,
        Title: "",
        DateofBirth: today.toISOString(),
        EmployeeType: "Select an option"
      },
      employeeTypes: []      
    };

    this._selection = new Selection({
      onSelectionChanged: this.onItemsSelectionChanged
    });
  }  

  private getEmployeeTypes(): Promise<IDropdownOption[]>{
    const url: string = this.props.siteUrl + "/_api/web/lists/GetByTitle('Employees')/fields?$filter=EntityPropertyName eq 'EmployeeType'";
    return this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
    .then(response => {
      return response.json();
    })
    .then(json => {      
      var _choices: IDropdownOption[]=[];      
      for(const r of json.value[0].Choices)
      {
        _choices.push({key: r, text: r});
      }      
      return _choices;
    }) as Promise<IDropdownOption[]>;
  }

  private _getListItems(): Promise<IListItem[]> {
    const url: string = this.props.siteUrl + "/_api/web/lists/getbytitle('Employees')/items";
    return this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
    .then(response => {
      return response.json();
    })
    .then(json => {
      return json.value;
    }) as Promise<IListItem[]>;
  }

  public bindDetailsList(message: string) : void {
    this._getListItems().then(_listItems => {
      this.setState({ listItems: _listItems, status: message });
    });
  }

  public componentDidMount(): void {
    this.getEmployeeTypes();
    this.bindDetailsList("All records loaded successfully");
  }

  @autobind
  public btnAdd_click(): void {
    
    const url: string = this.props.siteUrl + "/_api/web/lists/getbytitle('Demo')/items";

    const spHttpClientOptions: ISPHttpClientOptions = {
      "body": JSON.stringify(this.state.listItem)
    };

    this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
    .then((response: SPHttpClientResponse) => {
      if(response.status == 201){
        this.bindDetailsList("Record added and all records loaded successfully");
      } else {
        let errorMessage: string = "An error has occured " + response.status + " - " + response.statusText;
        this.setState({status: errorMessage});
      }
    });
  }

  @autobind
  public btnUpdate_click(): void {

    let id: number = this.state.listItem.Id;    
    const url: string = this.props.siteUrl + "/_api/web/lists/getbytitle('Demo')/items(" + id + ")";
    const headers: any = {
      "X-HTTP-Method": "MERGE",
      "IF-MATCH": "*"
    };

    const spHttpClientOptions: ISPHttpClientOptions = {
      "headers": headers,
      "body": JSON.stringify(this.state.listItem)
    };

    this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
    .then((response: SPHttpClientResponse) => {
      if(response.status == 204){
        this.bindDetailsList("Record updated and all records loaded successfully");
      } else {
        let errorMessage: string = "An error has occured " + response.status + " - " + response.statusText;
        this.setState({status: errorMessage});
      }
    });
  }

  @autobind
  public btnDelete_click(): void {

    let id: number = this.state.listItem.Id;
    const url: string = this.props.siteUrl + "/_api/web/lists/getbytitle('Demo')/items(" + id + ")";

    const headers: any = {
      "X-HTTP-Method": "DELETE",
      "IF-MATCH": "*"
    };

    const spHttpClientOptions: ISPHttpClientOptions = {
      "headers": headers      
    };

    this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
    .then((response: SPHttpClientResponse) => {
      if(response.status == 204){        
        this.bindDetailsList("Record deleted successfully");
      } else {
        let errorMessage: string = "An error has occured " + response.status + " - " + response.statusText;
        this.setState({status: errorMessage});
      }
    });
  }

  public render(): React.ReactElement<IReactSpFormProps> {

    const dropdownRef = React.createRef<IDropdown>();

    return (
      <div className={ styles.reactSpForm }>

        <TextField
          label="ID"
          required={ false }
          value={ (this.state.listItem.Id).toString() }
          className={controlClass.control}
          onChanged={e => {this.state.listItem.Id=e;}} />

        <TextField
          label="Full Name"
          required={ true }
          value={ this.state.listItem.Title }
          className={controlClass.control}
          onChanged={e => {this.state.listItem.Title=e;}} />

        <DatePicker
          label="Date of Birth"
          placeholder="Select date"
          className={controlClass.control}
          value={ new Date(this.state.listItem.DateofBirth) }
          onSelectDate={e => {this.state.listItem.DateofBirth=e.toISOString()}} />

        <Dropdown 
          componentRef={dropdownRef}
          placeholder="Select an option"
          label="Employee Type"
          options={this.state.employeeTypes}
          defaultSelectedKey={this.state.listItem.EmployeeType}
          required
          className={controlClass.control}
          onChanged={e => {this.state.listItem.EmployeeType=e.text;}} />

        <p className={styles.title}>      
            
            <PrimaryButton
              text='Add'
              title='Add'
              onClick={this.btnAdd_click} />
            
            <PrimaryButton
              text='Update'
              title='Update'
              onClick={this.btnUpdate_click} />

            <PrimaryButton
              text='Delete'
              title='Delete'
              onClick={this.btnDelete_click} /> 

        </p>

        <div id="divStatus">
          {this.state.status}
        </div>

        <div>
          <DetailsList
              items={ this.state.listItems }
              columns={_demoListColumns}
              setKey='Id'
              checkboxVisibility={CheckboxVisibility.always}
              selectionMode={SelectionMode.single}
              layoutMode={DetailsListLayoutMode.fixedColumns}
              compact={ true }
              selection={this._selection} />
        </div>

      </div>
    );
  }
}
