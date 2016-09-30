import * as React from 'react';

import {
  EnvironmentType,
  IHttpClientOptions,
  HttpClient
} from '@microsoft/sp-client-base';

import {
  css,
  FocusZone,
  FocusZoneDirection,
  List,
  ImageFit,
  Label,
  DocumentCard,
  DocumentCardActivity,
  DocumentCardTitle,
  IDocumentCardProps
} from 'office-ui-fabric-react';

import styles from '../CalendarReady.module.scss';
import { ICalendarReadyWebPartProps } from '../ICalendarReadyWebPartProps';
import { ICalendarItem } from '../services/CalendarService';
import { SPResult } from '../services/OfficeSearchService';



export interface ICalendarReadyProps extends ICalendarReadyWebPartProps {
}

export interface ICalendarState {
  items: ICalendarItem[],
  results: SPResult[]
}

export default class CalendarReady extends React.Component<ICalendarReadyProps, ICalendarState> {

  constructor(){
    super();

    this.state = {
      items: [] as ICalendarItem[],
      results: [] as SPResult[]
    };
  }

  public componentDidMount(): void {
    this._getCalendarItems();
  }

  public componentDidUpdate(): void {
    //Do nothing for now
  }

  public render(): JSX.Element {
    let { items, results } = this.state;
    return (
      <div>
        <FocusZone direction={ FocusZoneDirection.vertical} >
          <Label>Upcoming Meetings</Label>
          <List
            items= { items }
            onRenderCell={ (item, index) => (
              <div className='ms-ListBasicExample-itemCell'>
                  <div className='ms-ListBasicExample-itemContent'>
                    <div className='ms-ListBasicExample-itemName ms-font-xl'>{ item.Subject }</div>
                    <div className='ms-ListBasicExample-itemIndex'>{ `Item ${ index }` }</div>
                    <div className='ms-ListBasicExample-itemDesc ms-font-s'>{ item.Organizer.DisplayName }</div>
                  </div>
              </div>
            )}
          />
        </FocusZone>
        <FocusZone direction={ FocusZoneDirection.vertical }>
          <Label>Related Documents</Label>
          <List
            items={ results }
            onRenderCell={ (result) => (
              <div>{result.Title}</div>
            )} />
        </FocusZone>
      </div>
    );
  }

  private _getSearchResults(): void {
    this.props.httpClient.get(`${this.props.siteUrl}/_api/search/query?querytext='*'`, {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'odata-version': ''
      }
    })
    .then((response: Response) => {
    });
  }

  private _getCalendarItems(): void {
    this.props.httpClient.get('https://graph.microsoft.com/v1.0/me/events', this.graphHttpClientOptions)
    .then((response: Response) => {
      return response.json();
    })
    .then((jsonData: JSON) => {
      this._processCalendarItems(jsonData);
    });
  }

  private _processCalendarItems(jsonData: JSON): void {
    let calEntries: ICalendarItem[] = [];

    if(jsonData != null && jsonData["value"] != null && jsonData["value"].length > 0){

      let calItems: JSON[] = jsonData["value"];
      calItems.forEach((calItem: JSON) => {
        let calEntry: ICalendarItem = {
          StartDate: 0,
          EndDate: 0,
          Subject: calItem["subject"],
          Organizer: {
            DisplayName: "Paul Summers",
            EmailAddress: "paulsumm@microsoft.com",
            FirstName: "Paul",
            LastName: "Summers",
            ID: ""
          },
          People: [],
          Description: ""
        }

        calEntries.push(calEntry);
      });
    }
    this.setState({
        items: calEntries
    });
  }

  protected get graphHttpClientOptions(): IHttpClientOptions {
    return {
      headers: {
        'Authorization': 'Bearer eyJ0eXAiOiJKV1QiLCJub25jZSI6IkFRQUJBQUFBQUFEUk5ZUlEzZGhSU3JtLTRLLWFkcENKby1IV3NpeFU3RFhwcGxmX2ZNWlFoSXozc1loQ194SXc1TXh6Smc3b0lOUGVSV09OeXBZeXRDYmJQRjZhY3N5ZDBmcG9TZ1Y4V3hWSkdjM1dPNk1Rc2lBQSIsImFsZyI6IlJTMjU2IiwieDV0IjoiWWJSQVFSWWNFX21vdFdWSktIcndMQmJkXzlzIiwia2lkIjoiWWJSQVFSWWNFX21vdFdWSktIcndMQmJkXzlzIn0.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC80ZTlmNTQ2Ny01MjczLTQ3MGUtOWQwZi0wYTM5NmQ3YmZkNjcvIiwiaWF0IjoxNDc1MjA1MTI1LCJuYmYiOjE0NzUyMDUxMjUsImV4cCI6MTQ3NTIwOTAyNSwiYWNyIjoiMSIsImFtciI6WyJwd2QiXSwiYXBwX2Rpc3BsYXluYW1lIjoiU1BYMSIsImFwcGlkIjoiOGYyYTgwMDQtN2IxMi00ZjI0LThiNjAtNDRlZjhjMzAxOTk1IiwiYXBwaWRhY3IiOiIxIiwiZmFtaWx5X25hbWUiOiJTdW1tZXJzIiwiZ2l2ZW5fbmFtZSI6IlBhdWwiLCJpcGFkZHIiOiI3MS4yMzMuMTE1LjE2MSIsIm5hbWUiOiJQYXVsIFN1bW1lcnMiLCJvaWQiOiIyNDYyMTBjNy02NWJmLTRhOTUtYjBhZS1iMzAxZmM1Y2YwNzYiLCJwbGF0ZiI6IldpbiIsInB1aWQiOiIxMDAzN0ZGRTlBMDQ5NjdBIiwic2NwIjoiQ2FsZW5kYXJzLlJlYWQgQ2FsZW5kYXJzLlJlYWQuU2hhcmVkIENhbGVuZGFycy5SZWFkV3JpdGUgQ2FsZW5kYXJzLlJlYWRXcml0ZS5TaGFyZWQgQ29udGFjdHMuUmVhZCBDb250YWN0cy5SZWFkLlNoYXJlZCBDb250YWN0cy5SZWFkV3JpdGUgQ29udGFjdHMuUmVhZFdyaXRlLlNoYXJlZCBEaXJlY3RvcnkuQWNjZXNzQXNVc2VyLkFsbCBEaXJlY3RvcnkuUmVhZC5BbGwgRGlyZWN0b3J5LlJlYWRXcml0ZS5BbGwgZW1haWwgRmlsZXMuUmVhZCBGaWxlcy5SZWFkLkFsbCBGaWxlcy5SZWFkLlNlbGVjdGVkIEZpbGVzLlJlYWRXcml0ZSBGaWxlcy5SZWFkV3JpdGUuQWxsIEZpbGVzLlJlYWRXcml0ZS5BcHBGb2xkZXIgRmlsZXMuUmVhZFdyaXRlLlNlbGVjdGVkIEdyb3VwLlJlYWQuQWxsIEdyb3VwLlJlYWRXcml0ZS5BbGwgSWRlbnRpdHlSaXNrRXZlbnQuUmVhZC5BbGwgTWFpbC5SZWFkIE1haWwuUmVhZC5TaGFyZWQgTWFpbC5SZWFkV3JpdGUgTWFpbC5SZWFkV3JpdGUuU2hhcmVkIE1haWwuU2VuZCBNYWlsLlNlbmQuU2hhcmVkIE1haWxib3hTZXR0aW5ncy5SZWFkV3JpdGUgTm90ZXMuQ3JlYXRlIE5vdGVzLlJlYWQgTm90ZXMuUmVhZC5BbGwgTm90ZXMuUmVhZFdyaXRlIE5vdGVzLlJlYWRXcml0ZS5BbGwgTm90ZXMuUmVhZFdyaXRlLkNyZWF0ZWRCeUFwcCBvZmZsaW5lX2FjY2VzcyBvcGVuaWQgUGVvcGxlLlJlYWQgcHJvZmlsZSBTaXRlcy5SZWFkLkFsbCBUYXNrcy5SZWFkIFRhc2tzLlJlYWQuU2hhcmVkIFRhc2tzLlJlYWRXcml0ZSBUYXNrcy5SZWFkV3JpdGUuU2hhcmVkIFVzZXIuUmVhZCBVc2VyLlJlYWQuQWxsIFVzZXIuUmVhZEJhc2ljLkFsbCBVc2VyLlJlYWRXcml0ZSBVc2VyLlJlYWRXcml0ZS5BbGwiLCJzdWIiOiIyWUlPdnFpRWx0X2dNUzlGU0NpMlpxVGF2NzFmOW1BdmtnVnhPS1ItVGlnIiwidGlkIjoiNGU5ZjU0NjctNTI3My00NzBlLTlkMGYtMGEzOTZkN2JmZDY3IiwidW5pcXVlX25hbWUiOiJwYXVsQHBqc3VtbWVyc2pyZGV2Lm9ubWljcm9zb2Z0LmNvbSIsInVwbiI6InBhdWxAcGpzdW1tZXJzanJkZXYub25taWNyb3NvZnQuY29tIiwidmVyIjoiMS4wIn0.OtfzStOv-a3-zkcQ_aB-T1f-cE_SKLxdsJw2EGZZdkAJtOLjiOasfvYW4fYS6f3U6hNi5P6xXoTcJsbi_pBrAxgm0ZKOwvgfuY_erED6HVXFmjEGqvnev-iztCTV01uBw2XwNYsS0qXOQxBGrHrqWRTkvhcFFxqQCwp8g0q7YJ8gXk-2mPfQ4Di_cj4aSKneXi6Sd3jiKbMp33xdsbSKGW9sFKkPvAlQRYWfKA1ISbDUQU_mYZ7qsAkqObm0PidhOyLFJeWa-5fEQYIIzdKeESYEODSix2FLZSmtmOljq-eukEp3RXEsjuEmdUlKNvedNMP0HtB0Av-Ey_kKTBJUdw'
      }
    };
  }
}


