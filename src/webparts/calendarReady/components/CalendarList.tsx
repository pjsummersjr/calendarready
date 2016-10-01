import * as React from 'react';

import {
  IHttpClientOptions
} from '@microsoft/sp-client-base';

import { ICalendarItem } from '../services/CalendarService';
import { ICalendarReadyWebPartProps } from '../ICalendarReadyWebPartProps';

import {
  FocusZone,
  FocusZoneDirection,
  List,
  Label
} from 'office-ui-fabric-react';

export interface ICalendarListProps extends ICalendarReadyWebPartProps {
}

export interface ICalendarListState {
  items: ICalendarItem[]
}

export default class CalendarList extends React.Component<ICalendarListProps, ICalendarListState> {

  constructor(){
    super();

    this.state = {
      items: [] as ICalendarItem[]
    };
  }

  public componentDidMount(): void {
    this._getCalendarItems();
  }

  public render(): JSX.Element {
    let { items } = this.state;
    return (
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
    );
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
        'Authorization': 'Bearer eyJ0eXAiOiJKV1QiLCJub25jZSI6IkFRQUJBQUFBQUFEUk5ZUlEzZGhSU3JtLTRLLWFkcENKWm9CajI3S2E1WWZ1QmJKRzRJa3RJWGFCZDhSa2M5OE1xVWJHbndGSXIwUE9vYWl2YXdCOWFzUF9pclpkdmtEMnBqeWNQdGtDdmNTeFVTSnYzQzBrQ3lBQSIsImFsZyI6IlJTMjU2IiwieDV0IjoiWWJSQVFSWWNFX21vdFdWSktIcndMQmJkXzlzIiwia2lkIjoiWWJSQVFSWWNFX21vdFdWSktIcndMQmJkXzlzIn0.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC80ZTlmNTQ2Ny01MjczLTQ3MGUtOWQwZi0wYTM5NmQ3YmZkNjcvIiwiaWF0IjoxNDc1Mjg4MjA4LCJuYmYiOjE0NzUyODgyMDgsImV4cCI6MTQ3NTI5MjEwOCwiYWNyIjoiMSIsImFtciI6WyJwd2QiXSwiYXBwX2Rpc3BsYXluYW1lIjoiU1BYMSIsImFwcGlkIjoiOGYyYTgwMDQtN2IxMi00ZjI0LThiNjAtNDRlZjhjMzAxOTk1IiwiYXBwaWRhY3IiOiIxIiwiZmFtaWx5X25hbWUiOiJTdW1tZXJzIiwiZ2l2ZW5fbmFtZSI6IlBhdWwiLCJpcGFkZHIiOiIxNjcuMjIwLjE0OC4xOTMiLCJuYW1lIjoiUGF1bCBTdW1tZXJzIiwib2lkIjoiMjQ2MjEwYzctNjViZi00YTk1LWIwYWUtYjMwMWZjNWNmMDc2IiwicGxhdGYiOiJXaW4iLCJwdWlkIjoiMTAwMzdGRkU5QTA0OTY3QSIsInNjcCI6IkNhbGVuZGFycy5SZWFkIENhbGVuZGFycy5SZWFkLlNoYXJlZCBDYWxlbmRhcnMuUmVhZFdyaXRlIENhbGVuZGFycy5SZWFkV3JpdGUuU2hhcmVkIENvbnRhY3RzLlJlYWQgQ29udGFjdHMuUmVhZC5TaGFyZWQgQ29udGFjdHMuUmVhZFdyaXRlIENvbnRhY3RzLlJlYWRXcml0ZS5TaGFyZWQgRGlyZWN0b3J5LkFjY2Vzc0FzVXNlci5BbGwgRGlyZWN0b3J5LlJlYWQuQWxsIERpcmVjdG9yeS5SZWFkV3JpdGUuQWxsIGVtYWlsIEZpbGVzLlJlYWQgRmlsZXMuUmVhZC5BbGwgRmlsZXMuUmVhZC5TZWxlY3RlZCBGaWxlcy5SZWFkV3JpdGUgRmlsZXMuUmVhZFdyaXRlLkFsbCBGaWxlcy5SZWFkV3JpdGUuQXBwRm9sZGVyIEZpbGVzLlJlYWRXcml0ZS5TZWxlY3RlZCBHcm91cC5SZWFkLkFsbCBHcm91cC5SZWFkV3JpdGUuQWxsIElkZW50aXR5Umlza0V2ZW50LlJlYWQuQWxsIE1haWwuUmVhZCBNYWlsLlJlYWQuU2hhcmVkIE1haWwuUmVhZFdyaXRlIE1haWwuUmVhZFdyaXRlLlNoYXJlZCBNYWlsLlNlbmQgTWFpbC5TZW5kLlNoYXJlZCBNYWlsYm94U2V0dGluZ3MuUmVhZFdyaXRlIE5vdGVzLkNyZWF0ZSBOb3Rlcy5SZWFkIE5vdGVzLlJlYWQuQWxsIE5vdGVzLlJlYWRXcml0ZSBOb3Rlcy5SZWFkV3JpdGUuQWxsIE5vdGVzLlJlYWRXcml0ZS5DcmVhdGVkQnlBcHAgb2ZmbGluZV9hY2Nlc3Mgb3BlbmlkIFBlb3BsZS5SZWFkIHByb2ZpbGUgU2l0ZXMuUmVhZC5BbGwgVGFza3MuUmVhZCBUYXNrcy5SZWFkLlNoYXJlZCBUYXNrcy5SZWFkV3JpdGUgVGFza3MuUmVhZFdyaXRlLlNoYXJlZCBVc2VyLlJlYWQgVXNlci5SZWFkLkFsbCBVc2VyLlJlYWRCYXNpYy5BbGwgVXNlci5SZWFkV3JpdGUgVXNlci5SZWFkV3JpdGUuQWxsIiwic3ViIjoiMllJT3ZxaUVsdF9nTVM5RlNDaTJacVRhdjcxZjltQXZrZ1Z4T0tSLVRpZyIsInRpZCI6IjRlOWY1NDY3LTUyNzMtNDcwZS05ZDBmLTBhMzk2ZDdiZmQ2NyIsInVuaXF1ZV9uYW1lIjoicGF1bEBwanN1bW1lcnNqcmRldi5vbm1pY3Jvc29mdC5jb20iLCJ1cG4iOiJwYXVsQHBqc3VtbWVyc2pyZGV2Lm9ubWljcm9zb2Z0LmNvbSIsInZlciI6IjEuMCJ9.TZQ3_u9425X7NTF-8I0WQcANbp04zBL9LYDFjfJ-1OSyjBBSRZbeCQmEBRU8iat7JVp2oxckSQ_s8zS3WrH9yzhnXuEPKyVPTqgwJt0pk-JcbCRotLCvt9kXYcgqX1dKe2gtmsUEqq3636SS3oG-A9rn7o3WQ9p3AMRkYpfIce7y84iGclduJOdpgtS6u9sqR2CPkpG3Mc6lLwOoKRwBm8bS6iAqvO4HUiXbCf6BAOxmdkaSSFzOxNz7kc1lR8mI9CosWYBjNWJGXnuCPOiafbPAsOxXTxYrq1Ju-Wvi2a5nDFO4tvbx-RweHOYsP0W1V-1oq6-CLuBhW1hlQoyvjA'
      }
    };
  }
}