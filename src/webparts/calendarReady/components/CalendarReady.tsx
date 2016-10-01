import * as React from 'react';

import {
  EnvironmentType,
  IHttpClientOptions,
  HttpClient
} from '@microsoft/sp-client-base';

import styles from '../CalendarReady.module.scss';
import { ICalendarReadyWebPartProps } from '../ICalendarReadyWebPartProps';
import SearchResults, { ISearchResultsProps } from './SearchResults';
import CalendarList, { ICalendarListProps } from './CalendarList';



export interface ICalendarReadyProps extends ICalendarReadyWebPartProps {
}

export interface ICalendarReadyState {}

export default class CalendarReady extends React.Component<ICalendarReadyProps, ICalendarReadyState> {

  constructor(){
    super();
  }

  public componentDidMount(): void {
  }

  public componentDidUpdate(): void {
  }

  public render(): JSX.Element {
    return (
      <div>
        <CalendarList httpClient={ this.props.httpClient } siteUrl={ this.props.siteUrl } />
        <SearchResults httpClient={ this.props.httpClient } siteUrl={ this.props.siteUrl } />
      </div>
    );
  }
}


