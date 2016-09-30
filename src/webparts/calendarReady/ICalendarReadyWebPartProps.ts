import { HttpClient } from '@microsoft/sp-client-base';
import { ICalendarItem } from './services/CalendarService';

export interface IBasicWebPartProps {
  httpClient: HttpClient,
  siteUrl: string
}

export interface ICalendarReadyWebPartProps extends IBasicWebPartProps {
}

