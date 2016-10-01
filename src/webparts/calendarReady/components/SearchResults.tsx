import * as React from 'react';

import {
  EnvironmentType,
  IHttpClientOptions,
  HttpClient
} from '@microsoft/sp-client-base';

import { SharePointSearchClient, SPResults, SPResult } from '../services/OfficeSearchService';
import { ICalendarReadyWebPartProps} from '../ICalendarReadyWebPartProps';

import {
  FocusZone,
  FocusZoneDirection,
  Label,
  List
}
from 'office-ui-fabric-react';

export interface ISearchResultsProps extends ICalendarReadyWebPartProps {
}

export interface ISearchResultsState {
  results: SPResult[]
}

export default class SearchResults extends React.Component<ISearchResultsProps, ISearchResultsState> {
  constructor(){
    super();

    this.state = {
      results: [] as SPResult[]
    };
  }

  public componentDidMount(): void {
    this._getSearchResults();
  }

  public render(): JSX.Element {
    let { results } = this.state;
    return (

      <FocusZone direction={ FocusZoneDirection.vertical }>
        <Label>Related Documents</Label>
        <List
          items={ results }
          onRenderCell={ (result) => (
            <div>{result.Title}</div>
          )} />
      </FocusZone>
    );
  }

  private _getSearchResults(): void {

    let searchClient: SharePointSearchClient = new SharePointSearchClient();

    searchClient.Search(this.props.httpClient, this.props.siteUrl, "*")
    .then((response: SPResults) => {
      this.setState({
        results: response.value
      });
    });
  }
}