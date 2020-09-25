import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ReactSpFormWebPartStrings';
import ReactSpForm from './components/ReactSpForm';
import { IReactSpFormProps } from './components/IReactSpFormProps';

export interface IReactSpFormWebPartProps {
  listName: string;
}

export default class ReactSpFormWebPart extends BaseClientSideWebPart<IReactSpFormWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IReactSpFormProps> = React.createElement(
      ReactSpForm,
      {
        listName: this.properties.listName,
        context: this.context,
        siteUrl: this.context.pageContext.web.absoluteUrl

      }
    );

    ReactDom.render(element, this.domElement);
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
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
