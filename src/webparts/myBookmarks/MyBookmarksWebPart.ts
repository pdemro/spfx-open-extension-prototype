import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'MyBookmarksWebPartStrings';
import MyBookmarks from './components/MyBookmarks';
import { IMyBookmarksProps } from './components/IMyBookmarksProps';
import { MSGraphClient } from '@microsoft/sp-client-preview';

export interface IMyBookmarksWebPartProps {
  description: string;
}

export default class MyBookmarksWebPart extends BaseClientSideWebPart<IMyBookmarksWebPartProps> {

  public render(): void {

    const graphClient: MSGraphClient = this.context.serviceScope.consume(
      MSGraphClient.serviceKey
    )

    graphClient
      .api("me")
      .version("v1.0")
      .select("id,displayName")
      .expand("extensions")
      .get((err, res) => {
        if(err) {
          console.error(err);
          return;
        }

        console.log(res);
      })


    const element: React.ReactElement<IMyBookmarksProps > = React.createElement(
      MyBookmarks,
      {
        description: this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);
  }

  // private test (): void {
  //   const graphClient: MSGraphClient = this.context.serviceScope.
  // }

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
