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
        let hour = date.getDate();
        let updatedDate = (date.toString()).slice(0,15);
        html += `
          <li class=${styles.liEV}>
            <p class=${styles.dateHeaderEV}>${updatedDate} <a href=#>>></a></p>
            <p class=${styles.eventEV}>{start time} ${item.Title}</p>
            <p class=${styles.subEventEV}>{event length} ${item.Location}</p>
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
