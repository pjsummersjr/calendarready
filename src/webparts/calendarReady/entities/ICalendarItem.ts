export interface ICalendarItem {
  StartDate: number;
  EndDate: number;
  Subject: string;
  Description: string;
  Organizer: Person;
  People: Person[];
}

export interface Person {
  DisplayName: string;
  FirstName: string;
  LastName: string;
  EmailAddress: string;
  ID: string;
}