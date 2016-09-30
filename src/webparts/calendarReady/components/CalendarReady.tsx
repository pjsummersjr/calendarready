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
import { SharePointSearchClient, SPResults, SPResult } from '../services/OfficeSearchService';



export interface ICalendarReadyProps extends ICalendarReadyWebPartProps {
}

export interface ICalendarState {
  items: ICalendarItem[],
  results?: SPResult[]
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
    this._getSearchResults();
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

    let searchClient: SharePointSearchClient = new SharePointSearchClient();

    searchClient.Search(this.props.httpClient, this.props.siteUrl, "*")
    .then((response: SPResults) => {
      this.setState({
        items: this.state.items,
        results: response.value
      });
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
        items: calEntries,
        results: this.state.results
    });
  }

  protected get graphHttpClientOptions(): IHttpClientOptions {
    return {
      headers: {
        'Authorization': 'Bearer eyJ0eXAiOiJKV1QiLCJub25jZSI6IkFRQUJBQUFBQUFEUk5ZUlEzZGhSU3JtLTRLLWFkcENKZElfZVVwTlA2NzBkaTVkNENIS1k1VEZPV2tXLUtCUVN6RVhia2hlOEljOXhldHFFdHpfd0tSd01GN3BjRThTLUZXV1NmSzZ4cm5XLXpGOWg5dDdsVGlBQSIsImFsZyI6IlJTMjU2IiwieDV0IjoiSTZvQnc0VnpCSE9xbGVHclYyQUpkQTVFbVhjIiwia2lkIjoiSTZvQnc0VnpCSE9xbGVHclYyQUpkQTVFbVhjIn0.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC80ZTlmNTQ2Ny01MjczLTQ3MGUtOWQwZi0wYTM5NmQ3YmZkNjcvIiwiaWF0IjoxNDc1MjYxOTA2LCJuYmYiOjE0NzUyNjE5MDYsImV4cCI6MTQ3NTI2NTgwNiwiYWNyIjoiMSIsImFtciI6WyJwd2QiXSwiYXBwX2Rpc3BsYXluYW1lIjoiU1BYMSIsImFwcGlkIjoiOGYyYTgwMDQtN2IxMi00ZjI0LThiNjAtNDRlZjhjMzAxOTk1IiwiYXBwaWRhY3IiOiIxIiwiZmFtaWx5X25hbWUiOiJTdW1tZXJzIiwiZ2l2ZW5fbmFtZSI6IlBhdWwiLCJpcGFkZHIiOiIxNjcuMjIwLjE0OC4xOTMiLCJuYW1lIjoiUGF1bCBTdW1tZXJzIiwib2lkIjoiMjQ2MjEwYzctNjViZi00YTk1LWIwYWUtYjMwMWZjNWNmMDc2IiwicGxhdGYiOiJXaW4iLCJwdWlkIjoiMTAwMzdGRkU5QTA0OTY3QSIsInNjcCI6IkNhbGVuZGFycy5SZWFkIENhbGVuZGFycy5SZWFkLlNoYXJlZCBDYWxlbmRhcnMuUmVhZFdyaXRlIENhbGVuZGFycy5SZWFkV3JpdGUuU2hhcmVkIENvbnRhY3RzLlJlYWQgQ29udGFjdHMuUmVhZC5TaGFyZWQgQ29udGFjdHMuUmVhZFdyaXRlIENvbnRhY3RzLlJlYWRXcml0ZS5TaGFyZWQgRGlyZWN0b3J5LkFjY2Vzc0FzVXNlci5BbGwgRGlyZWN0b3J5LlJlYWQuQWxsIERpcmVjdG9yeS5SZWFkV3JpdGUuQWxsIGVtYWlsIEZpbGVzLlJlYWQgRmlsZXMuUmVhZC5BbGwgRmlsZXMuUmVhZC5TZWxlY3RlZCBGaWxlcy5SZWFkV3JpdGUgRmlsZXMuUmVhZFdyaXRlLkFsbCBGaWxlcy5SZWFkV3JpdGUuQXBwRm9sZGVyIEZpbGVzLlJlYWRXcml0ZS5TZWxlY3RlZCBHcm91cC5SZWFkLkFsbCBHcm91cC5SZWFkV3JpdGUuQWxsIElkZW50aXR5Umlza0V2ZW50LlJlYWQuQWxsIE1haWwuUmVhZCBNYWlsLlJlYWQuU2hhcmVkIE1haWwuUmVhZFdyaXRlIE1haWwuUmVhZFdyaXRlLlNoYXJlZCBNYWlsLlNlbmQgTWFpbC5TZW5kLlNoYXJlZCBNYWlsYm94U2V0dGluZ3MuUmVhZFdyaXRlIE5vdGVzLkNyZWF0ZSBOb3Rlcy5SZWFkIE5vdGVzLlJlYWQuQWxsIE5vdGVzLlJlYWRXcml0ZSBOb3Rlcy5SZWFkV3JpdGUuQWxsIE5vdGVzLlJlYWRXcml0ZS5DcmVhdGVkQnlBcHAgb2ZmbGluZV9hY2Nlc3Mgb3BlbmlkIFBlb3BsZS5SZWFkIHByb2ZpbGUgU2l0ZXMuUmVhZC5BbGwgVGFza3MuUmVhZCBUYXNrcy5SZWFkLlNoYXJlZCBUYXNrcy5SZWFkV3JpdGUgVGFza3MuUmVhZFdyaXRlLlNoYXJlZCBVc2VyLlJlYWQgVXNlci5SZWFkLkFsbCBVc2VyLlJlYWRCYXNpYy5BbGwgVXNlci5SZWFkV3JpdGUgVXNlci5SZWFkV3JpdGUuQWxsIiwic3ViIjoiMllJT3ZxaUVsdF9nTVM5RlNDaTJacVRhdjcxZjltQXZrZ1Z4T0tSLVRpZyIsInRpZCI6IjRlOWY1NDY3LTUyNzMtNDcwZS05ZDBmLTBhMzk2ZDdiZmQ2NyIsInVuaXF1ZV9uYW1lIjoicGF1bEBwanN1bW1lcnNqcmRldi5vbm1pY3Jvc29mdC5jb20iLCJ1cG4iOiJwYXVsQHBqc3VtbWVyc2pyZGV2Lm9ubWljcm9zb2Z0LmNvbSIsInZlciI6IjEuMCJ9.BOvaXF-w8Kp5TO-yJ3fJzKt4DgaMPAvD5Hf5hbt9C4oJMEyLQjV422cFXiMfUxBFLKEjjfz4gzdAtGNjYwPgNB_khRm1XEXSZW9sMLDhml_D7M7qlYyaNKD8CfQ0WhdkFIrdk0Pjljkh4u49XHRlSMxNfwKWUFepY5x5FD5CcFzIVgd1frVFupqQWAejq9cThwUkT8VVFZ9G9GJ5qh_afcd7YrgW_6FtgNYqIEeVmx-y569RUS0nb9Joywxg2k9aTObQKPVqzHjLdshCrFRZx_aTa_NZkMiuPoH_CC4Eb2ksVtMZGvtiNKo_XxcmRPHvrApNdOAeHQae-GcIhcMFzA'
      }
    };
  }
}


