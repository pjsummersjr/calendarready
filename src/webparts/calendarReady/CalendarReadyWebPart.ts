import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneTextField
} from '@microsoft/sp-client-preview';

import * as strings from 'calendarReadyStrings';
import CalendarReady, { ICalendarReadyProps } from './components/CalendarReady';
import { ICalendarReadyWebPartProps } from './ICalendarReadyWebPartProps';

export default class CalendarReadyWebPart extends BaseClientSideWebPart<ICalendarReadyWebPartProps> {

  public constructor(context: IWebPartContext) {
    super(context);
  }

  public render(): void {
    const element: React.ReactElement<ICalendarReadyProps> = React.createElement(CalendarReady, {
      httpClient: this.context.httpClient,
      siteUrl: this.context.pageContext.web.absoluteUrl
      //description: this.properties.description
    });

    ReactDom.render(element, this.domElement);
  }

  protected get propertyPaneSettings(): IPropertyPaneSettings {
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
