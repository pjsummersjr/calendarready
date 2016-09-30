import { HttpClient, IHttpClientOptions } from '@microsoft/sp-client-base';

export interface SPResults {
  value: SPResult[];
}

export interface SPResult {
  Id: string;
  Title: string;
  Url: string;
}

export interface SharePointResults {
  ElapsedTime: number;
  PrimaryQueryResult: Object[];
}

export class SharePointSearchClient {

  public Search(httpClient: HttpClient, siteUrl: string, queryText: string): Promise<SPResults> {
    let resultData: SPResults = { value: [] };
    return this._getSearchResults(httpClient, siteUrl, queryText).then((response) => {
      return this._processResults(response)});
  }

  private _getSearchResults(httpClient: HttpClient, siteUrl: string, queryText: string): Promise<JSON> {
    const httpOptions: IHttpClientOptions = this.searchHttpClientOptions;

    return httpClient.get(siteUrl + `/_api/search/query?querytext=%27` + queryText + `%27`, httpOptions)
    .then((response: Response) => {
      return response.json();
    });
  }

  private _processResults(response: JSON): SPResults{
    if(response != null && response['PrimaryQueryResult'] != null &&
      response['PrimaryQueryResult']['RelevantResults'] != null &&
      response['PrimaryQueryResult']['RelevantResults']['Table'] != null &&
      response['PrimaryQueryResult']['RelevantResults']['Table']['Rows'] != null){

      let resultData: SPResults = { value: [] };

      let rows = response['PrimaryQueryResult']['RelevantResults']['Table']['Rows'];
      rows.forEach((row: JSON) => {
        let result: SPResult = {Id: '', Title: '', Url: ''};
        let cells = row['Cells'];
        cells.forEach((cell: JSON) => {
          if(cell['Key'] == 'DocId'){
            result.Id = cell['Value'];
          }
          if(cell['Key'] == 'Title'){
            result.Title = cell['Value'];
          }
          if(cell['Key'] == 'Path'){
            result.Url = cell['Value'];
          }
        });
        resultData.value.push(result);
      });
      return resultData;
    }
  }

  protected get searchHttpClientOptions(): IHttpClientOptions {
    return {
      headers: {
        'odata-version': ''
      }
    };
  }

}