import { HttpClient } from '@microsoft/sp-client-base';
import { ICalendarItem } from './entities/ICalendarItem';

export interface IBasicWebPartProps {
  httpClient: HttpClient
}

export interface ICalendarReadyWebPartProps extends IBasicWebPartProps {
  items: ICalendarItem[];
}

