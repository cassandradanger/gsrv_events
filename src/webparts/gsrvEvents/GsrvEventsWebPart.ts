import { Version } from '@microsoft/sp-core-library';
import { sp, Items, ItemVersion, Web } from "@pnp/sp";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import {  
  SPHttpClient,
  SPHttpClientResponse   
} from '@microsoft/sp-http';

import styles from './GsrvEventsWebPart.module.scss';
import * as strings from 'GsrvEventsWebPartStrings';

import * as moment from 'moment';
import { dateAdd } from '@pnp/common';
let now = moment().format('LLLL');

export interface IGsrvEventsWebPartProps {
  description: string;
}
export interface ISPLists {
  value: ISPList[];
 }

 export interface ISPList {
  Title: string; // this is the department name in the List
  Id: string;
  AnncURL:string;
  DeptURL:string;
  CalURL:string;
  a85u:string; // this is the LINK URL
 }

export interface IGsvrDeptEventsWebPartProps {
  description: string;
}

//global vars
var userDept = "";

export default class GsrvEventsWebPart extends BaseClientSideWebPart<IGsrvEventsWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <div class=${styles.mainEV}>
    <ul class=${styles.contentEV}>
        <div id="spListContainer" /></div>
        </ul>
    </div>`;
  }

  getuser = new Promise((resolve,reject) => {
    // SharePoint PnP Rest Call to get the User Profile Properties
    return sp.profiles.myProperties.get().then(function(result) {
      var props = result.UserProfileProperties;
      props.forEach(function(prop) {
        //this call returns key/value pairs so we need to look for the Dept Key
        if(prop.Key == "Department"){
          // set our global var for the users Dept.
          userDept += prop.Value;
        }
      });
      return result;
    }).then((result) =>{
      this._getListData().then((response) =>{
        this._renderList(response.value);
      });
    });
  });

  public _getListData(): Promise<ISPLists> {  
    return this.context.spHttpClient.get(`https://girlscoutsrv.sharepoint.com/_api/web/lists/GetByTitle('TeamDashboardSettings')/Items?$filter=Title eq '`+ userDept +`'`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
   }

   private _renderList(items: ISPList[]): void {
    let html: string = ``;
      html += `
        <h3 class=${styles.titleEV}>
          <a href="https://girlscoutsrv.sharepoint.com${items[0].DeptURL}/Lists/Team%20Events">Upcoming Deadlines and Team Calendar</a>
        </h3>
        `
    var siteURL = "";
    var eventsListName =  "";
    var date = new Date();
    var strToday = "";
    var mm = date.getMonth()+1;
    var dd = date.getDate();
    var yyyy = date.getFullYear();
    if(dd < 10){
      dd = 0 + dd;
    }
    if(mm < 10){
      mm = 0 + mm;
    }
    strToday = mm + "/" + dd + "/" + yyyy;

    items.forEach((item: ISPList) => {
      siteURL = item.DeptURL;
      eventsListName = item.CalURL;
    });
    //1st we need to override the current web to go to the department sites web
    const w = new Web("https://girlscoutsrv.sharepoint.com" + siteURL);
    
    // then use PnP to query the list
    // CASIE IF YOU NEED MORE THAN 5 EVENTS JUST UPDATE THE NUMBER BELOW
    w.lists.getByTitle(eventsListName).items.filter("EventDate ge '" + strToday + "'").top(3).orderBy("EventDate")
    .get()
    .then((data) => {
      let date: any;
      let endDate: any;
      let dayName = '';
      let monthName = '';
      let dayNum = '';
      let year = '';
      let duration: any;
      let endTime: any;

      data.forEach((data) => {
        if(data.fAllDayEvent){
          date = new Date(data.EventDate).toUTCString();
          console.log(date);
          endDate = new Date(data.EndDate).toUTCString();
          dayName = (date.toString()).slice(0,3);
          monthName = (date.toString()).slice(8,11);
          dayNum = (date.toString()).slice(4,7);
          year = (date.toString()).slice(12,16);
        }else {
        date = new Date(data.EventDate);
        endDate = new Date(data.EndDate);
        

        console.log(date);
        duration = moment(endDate,"DD/MM/YYYY HH:mm:ss").diff(moment(date,"DD/MM/YYYY HH:mm:ss"));
        endTime = moment.duration(duration).asMinutes();

        dayName = (date.toString()).slice(0,3);
        monthName = (date.toString()).slice(4,7);
        dayNum = (date.toString()).slice(8,10);
        year = (date.toString()).slice(11,15);
        }

        let startTime = (date.toString()).slice(16, 24);
        let standardStartTime = moment(startTime, 'HH:mm').format('hh:mm a');
        let location = data.Location;

        let displayTime = '0';
        let displayType = '';

        if(data.fAllDayEvent){
          displayTime = 'All Day';
          displayType = ''
        } else if(endTime === 60){
          displayTime = '1';
          displayType = 'hour'
        } else if (endTime >= 60 && endTime <= 1440){
          let tempTime = Math.ceil(endTime/60);
          displayTime = tempTime.toString();
          displayType = 'hours';
        } else if (endTime > 1440){
          let tempTime = Math.ceil(endTime / 1440);
          displayTime = tempTime.toString();
          displayType = 'days'
        }
        else if (endTime < 60){
          displayTime = endTime.toString();
          displayType = endTime > 1 ? 'minutes' : 'minute';
        }

        if(location === null){
          location = "Location TBD";
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
            <p class=${styles.dateHeaderEV}>${dayName}, ${monthName} ${dayNum}, ${year} 
              <a href="https://girlscoutsrv.sharepoint.com${siteURL}/Lists/${eventsListName}/DispForm.aspx?ID=${data.Id}">>></a>
            </p>
            <p class=${styles.eventEV}>${standardStartTime} ${data.Title}</p>
            <p class=${styles.subEventEV}>${displayTime} ${displayType} <span class=${styles.locationEV}>${location}</span></p>
            <div class=${styles.verticalBar}></div>
          </li>
          `;  
      })
      var secondHtml = `<p class=${styles.gsrv_logo}></p>`
      if(items.length === 0){
        const listContainer: Element = this.domElement.querySelector('#spListContainer');  
        listContainer.innerHTML = secondHtml;  
      } else {
        const listContainer: Element = this.domElement.querySelector('#spListContainer');  
        listContainer.innerHTML = html;  
      }
    }).catch(e => { console.error(e); });
    
  }

    // this is required to use the SharePoint PnP shorthand REST CALLS
    public onInit():Promise<void> {
      return super.onInit().then (_=> {
        sp.setup({
          spfxContext:this.context
        });
      });
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
