import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './NewsSliderWebPart.module.scss';
import * as strings from 'NewsSliderWebPartStrings';
import {  
  SPHttpClient  
} from '@microsoft/sp-http';  

export interface ISPLists {
  value: ISPList[];  
}

export interface ISPList{
  ShowOnSlider: Boolean;
  SliderImage: any;
}

export interface INewsSliderWebPartProps {
  description: string;
}

export default class NewsSliderWebPart extends BaseClientSideWebPart<INewsSliderWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <div class=${styles.mainNS}>
       TESTING
       <div id="spListContainer" /></div>
    </div>`;
      this._firstGetList();
  }
  private _firstGetList() {
    this.context.spHttpClient.get('https://girlscoutsrv.sharepoint.com' + 
      `/_api/web/Lists/GetByTitle('Stories, News, & Announcements')/Items?)`, SPHttpClient.configurations.v1)
      .then((response)=>{
        response.json().then((data)=>{
          console.log('this is data', data);
          this._renderList(data.value)
        })
      });
  }

  private _renderList(items: ISPList[]): void {
    let html: string = ``;
    let image: any;
    items.forEach((item: ISPList) => {
      if(item.ShowOnSlider !== true){
        return;
      }else{
        image = item.SliderImage.Url;
        console.log(image);
      }
      html += ` 
      <p>here's an image</p>
      <img src="${image}"/>
      `
    });
    const listContainer: Element = this.domElement.querySelector('#spListContainer');  
    listContainer.innerHTML = html;
  };

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
