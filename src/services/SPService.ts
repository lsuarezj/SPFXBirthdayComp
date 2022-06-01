import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import * as moment from "moment";

export class SPService {
  // private graphClient: MSGraphClient = null;
  private birthdayListTitle: string = "Birthdays";
  constructor(private _context: WebPartContext) {}
  // Get Profiles
  public async getPBirthdays(
    upcommingDays: number
  ): Promise<SPHttpClientResponse> {
    let _results, _today: string, _month: number, _day: number;
    let _filter: string, _countdays: number, _f: number, _nextYearStart: string;
    let _FinalDate: string;
    try {
      _results = null;
      _today = "2000-" + moment().format("MM-DD");
      _month = parseInt(moment().format("MM"));
      _day = parseInt(moment().format("DD"));
      _filter = "Birthday ge '" + _today + "'";
      // If we are in December we have to look if there are birthdays in January
      // we have to build a condition to select birthday in January based on number of upcommingDays
      // we can not use the year for test, the year is always 2000.

      // _countdays = _day + upcommingDays;
      _countdays = _day + 30;
      _f = 0;
      if (_month === 12 && _countdays > 31) {
        _nextYearStart = "2000-01-01";
        _FinalDate = "2000-01-";
        _f = _countdays - 31;
        _FinalDate = _FinalDate + _f;
        _filter =
          "Birthday ge '" +
          _today +
          "' or (Birthday ge '" +
          _nextYearStart +
          "' and Birthday le '" +
          _FinalDate +
          "')";
      } else {
        _FinalDate = "2000-";
        if (_countdays > 31) {
          _f = _countdays - 31;
          _month = _month + 1;
          _FinalDate = _FinalDate + _month + "-" + _f;
        } else {
          _FinalDate = _FinalDate + _month + "-" + _countdays;
        }
        _filter =
          "Birthday ge '" + _today + "' and Birthday le '" + _FinalDate + "'";
      }

      let apiUrl =
        this._context.pageContext.web.absoluteUrl +
        `/_api/web/lists/GetByTitle('${this.birthdayListTitle}')/Items?$filter=${_filter}&$top=${upcommingDays}&$orderBy=Birthday`;

      return this._context.spHttpClient
        .get(apiUrl, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
          return response.json();
        });
    } catch (error) {
      console.dir(error);
      return Promise.reject(error);
    }
  }
}
export default SPService;
