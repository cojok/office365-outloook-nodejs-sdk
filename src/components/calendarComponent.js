import { findIana } from 'windows-iana';
import moment from 'moment-timezone';
export default class CalendarComponent {
  constructor(authService) {
    this.authService = authService;
  }
  async getAllEvents(accessToken, timeZone) {
    if (accessToken) {
      try {
        const client = this.authService.getAuthenticatedClient(accessToken);
        const timeZoneId = findIana(timeZone)[0];
        const startOfWeek = moment.tz(timeZoneId.valueOf()).startOf('week').utc();
        const endOfWeek = moment(startOfWeek).subtract(45, 'day');
        const viewData = await client.api('/me/events')
        .header('Prefer', `outlook.timezone="${timeZone}"`)
        .query({ startDateTime: endOfWeek.format(), endDateTime: startOfWeek.format() })
        .select('subject,organizer,start,end,attendees')
        .orderby('start/dateTime')
        .get();
        return viewData.value.map((data) => {
          delete data['@odata.etag'];
          return data;
        });
      } catch (error) {
        console.log(JSON.stringify(error, Object.getOwnPropertyNames(error)));
        throw new Error(error);
      }
    } else {
      throw new Error('No access token provided');
    }
  }

  // TODO: decide which elements should be selected to be returned by the endpoint
  async getEventById(accessToken, timeZone, id) {
    if(accessToken) {
      try {
        const client = this.authService.getAuthenticatedClient(accessToken);
        const timeZoneId = findIana(timeZone)[0];
        // const startOfWeek = moment.tz(timeZoneId.valueOf()).startOf('week').utc();
        // const endOfWeek = moment(startOfWeek).subtract(45, 'day');
        const viewData = await client.api(`/me/events/${id}`)
        .header('Prefer', `outlook.timezone="${timeZone}"`)
        // .query({ startDateTime: endOfWeek.format(), endDateTime: startOfWeek.format() })
        // .select('subject,organizer,start,end,attendees')
        // .orderby('start/dateTime')
        .get();
        delete viewData['@odata.context'];
        delete viewData['@odata.etag'];
        return viewData;
      } catch (error) {
        console.log(JSON.stringify(error, Object.getOwnPropertyNames(error)));
        throw new Error(error);
      }
    } else {
      throw new Error('No access token provided');
    }   
  }
}