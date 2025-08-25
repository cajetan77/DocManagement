import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';

import * as strings from 'DocumentLibraryViewsWebPartStrings';
import DocumentLibraryViews from './components/DocumentLibraryViews';
import { IDocumentLibraryViewsProps } from './components/IDocumentLibraryViewsProps';

export interface IDocumentLibraryViewsWebPartProps {
  description: string;
  siteUrl: string;
  listTitle: string;
}

export default class DocumentLibraryViewsWebPart extends BaseClientSideWebPart<IDocumentLibraryViewsWebPartProps> {

  public render(): void {
    // Debug logging to see what properties are being used
    console.log('=== WebPart Properties Debug ===');
    console.log('Properties siteUrl:', this.properties.siteUrl);
    console.log('Properties listTitle:', this.properties.listTitle);
    console.log('Current site URL:', this.context.pageContext.web.absoluteUrl);
    console.log('Final siteUrl:', this.properties.siteUrl || this.context.pageContext.web.absoluteUrl);
    console.log('Final listTitle:', this.properties.listTitle || '');

    const element: React.ReactElement<IDocumentLibraryViewsProps> = React.createElement(
      DocumentLibraryViews,
      {
        description: this.properties.description,
        isDarkTheme: false,
        environmentMessage: this._getEnvironmentMessage(),
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        siteUrl: this.properties.siteUrl || this.context.pageContext.web.absoluteUrl,
        listTitle: this.properties.listTitle || 'Working Document Library',
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse(this.context.manifest.version);
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
                }),
                PropertyPaneTextField('siteUrl', {
                  label: 'Site URL',
                  description: 'The URL of the site containing the document library (leave empty to use current site)',
                  placeholder: 'https://yourtenant.sharepoint.com/sites/yoursite'
                }),
                PropertyPaneTextField('listTitle', {
                  label: 'Document Library Title',
                  description: 'The title of the document library (defaults to "Working Document Library")',
                  placeholder: 'Working Document'
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) {
      return strings.AppTeamsTabEnvironment;
    }

    return strings.AppSharePointEnvironment;
  }
}
