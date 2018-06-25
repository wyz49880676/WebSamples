import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCustomField
} from '@microsoft/sp-webpart-base';

import * as strings from 'RecentDocsWebPartStrings';
import RecentDocs from './components/Folders';
import { IRecentDocsProps,IRecentDocsWebPartProps } from './Model/IRecentDocsProps';

export default class RecentDocsWebPart extends BaseClientSideWebPart<IRecentDocsWebPartProps> {

  public render(): void {

    const element: React.ReactElement<IRecentDocsProps> = React.createElement(
      RecentDocs,
      {
        context: this.context,
        listUrl: this.properties.listUrl,
        listTitle: this.properties.listTitle,
        siteUrl: this.properties.siteUrl
      }
    );

    ReactDom.render(element, this.domElement);
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
              groupName: strings.InfoGroupName,
              groupFields: [
                PropertyPaneTextField('siteUrl', {
                  label: strings.SiteURLLabel,
                  value: this.context.pageContext.web.absoluteUrl
                }),
                PropertyPaneTextField('listTitle', {
                  label: strings.ListNameLabel,
                  value: 'Documents'
                }),
                PropertyPaneTextField('listUrl', {
                  label: strings.ListURLLabel,
                  value: 'Shared Documents'
                })
              ]
            },
            // {
            //   groupName: strings.DataGroupName,
            //   groupFields: [
            //     PropertyPaneTextField('siteUrl', {
            //       label: strings.SiteURLLabel,
            //       value: this.context.pageContext.web.absoluteUrl
            //     }),
            //     PropertyPaneTextField('listTitle', {
            //       label: strings.ListNameLabel,
            //       value: 'Documents'
            //     }),
            //     PropertyPaneTextField('listUrl', {
            //       label: strings.ListURLLabel,
            //       value: 'Shared Documents'
            //     })
            //   ]
            // }
          ]
        }
      ]
    };
  }
}
