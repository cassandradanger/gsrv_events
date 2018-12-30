import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import {  
  SPHttpClient  
} from '@microsoft/sp-http';

import styles from './GsrvEventsWebPart.module.scss';
import * as strings from 'GsrvEventsWebPartStrings';

import * as moment from 'moment';
let now = moment().format('LLLL');

export interface IGsrvEventsWebPartProps {
  description: string;
}
export interface ISPLists {
  value: ISPList[];  
}

export interface ISPList{
  EventDate: any;
  Title: any;
  Location: any;
  EndDate: any;
}


export default class GsrvEventsWebPart extends BaseClientSideWebPart<IGsrvEventsWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <div class=${styles.mainEV}>
      <p class=${styles.titleEV}>
        Upcoming Deadlines and Team Calendar
      </p>
      <ul class=${styles.contentEV}>
        <div id="spListContainer" /></div>
      </ul>
    </div>`;
      this._firstGetList();
  }

  private _firstGetList() {
    this.context.spHttpClient.get(`https://girlscoutsrv.sharepoint.com/gsrv_teams/sdgdev/_api/web/lists/getbytitle('Events')/items()`, SPHttpClient.configurations.v1)
      .then((response)=>{
        response.json().then((data)=>{
          console.log(data.value);
          this._renderList(data.value)
        })
      });
    }
  


    private _renderList(items: ISPList[]): void {
      let html: string = ``;
      
      items.forEach((item: ISPList) => {
        let date = new Date(item.EventDate);
        let endDate = new Date(item.EndDate);
        let testing = moment(endDate,"DD/MM/YYYY HH:mm:ss").diff(moment(date,"DD/MM/YYYY HH:mm:ss"));
        let endTime = moment.duration(testing);

        let dayName = (date.toString()).slice(0,3);
        let monthName = (date.toString()).slice(4,7);
        let dayNum = (date.toString()).slice(8, 9);
        let year = (date.toString()).slice(10,15);

        let startTime = (date.toString()).slice(16, 24);
        let standardStartTime = moment(startTime, 'HH:mm').format('hh:mm a');
        let location = item.Location;

        let displayHour = endTime.hours().toString();
        let displayMinute = endTime.minutes().toString();
        let displayTime = '';

        if(location === null){
          location = "Location TBD";
        }
        if(displayHour === '0'){
          displayHour = '';
        }
        if(displayMinute === '0'){
          displayMinute = '';
        }

        if(endTime.hours() === 1){
          displayTime = 'hour'
        } else if (endTime.hours() > 1){
          displayTime = 'hours'
        } else if (endTime.minutes() > 0){
          displayTime  = 'minutes'
        } else if(endTime.hours() === 0 && endTime.minutes() === 0){
          displayTime = 'All day'
        }

        if(dayName === 'Mon'){
          dayName = 'Monday';
        } else if (dayName ==='Tue'){
          dayName = 'Tuesday';
        } else if (dayName ==='Wed'){
          dayName = 'Wednesday';
        } else if (dayName ==='Thu'){
          dayName = 'Thursday';
        } else if (dayName ==='Fri'){
          dayName = 'Friday';
        } else if (dayName ==='Sat'){
          dayName = 'Saturday';
        } else if (dayName ==='Sun'){
          dayName = 'Sunday';
        } 

        if(monthName === 'Jan'){
          monthName = 'January';
        } else if (monthName ==='Feb'){
          monthName = 'February';
        } else if (monthName ==='Mar'){
          monthName = 'March';
        } else if (monthName ==='Apr'){
          monthName = 'April';
        } else if (monthName ==='May'){
          monthName = 'May';
        } else if (monthName ==='Jun'){
          monthName = 'June';
        } else if (monthName ==='Jul'){
          monthName = 'Jul';
        } else if (monthName ==='Aug'){
          monthName = 'August';
        } else if (monthName ==='Sep'){
          monthName = 'September';
        } else if (monthName ==='Oct'){
          monthName = 'October';
        } else if (monthName ==='Nov'){
          monthName = 'November';
        } else if (monthName ==='Dec'){
          monthName = 'December';
        } 
        html += `
          <li class=${styles.liEV}>
            <p class=${styles.dateHeaderEV}>${dayName}, ${monthName} ${dayNum}, ${year} <a href=#>>></a></p>
            <p class=${styles.eventEV}>${standardStartTime} ${item.Title}</p>
            <p class=${styles.subEventEV}>${displayHour} ${displayMinute} ${displayTime} <span class=${styles.locationEV}>${location}</span></p>
          </li>
          `;  
      });  
      const listContainer: Element = this.domElement.querySelector('#spListContainer');  
      listContainer.innerHTML = html;  
    } 
  
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
