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
  Title: String;
}

export interface INewsSliderWebPartProps {
  description: string;
}

export default class NewsSliderWebPart extends BaseClientSideWebPart<INewsSliderWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <div class=${styles.headerNS}>
    <p class=${styles.announceNS}></p>
       News & Announcements
       <div id="buttons" / class=${styles.buttonsWrap}></div>
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
    let buttons: string = ``;
    let objectArray: Array<any> = [];
    let buttonNumber = 0;
    let imageNumber = 0;
    items.forEach((item: ISPList) => {
      if(item.ShowOnSlider !== true){
        return;
      }else{
        buttonNumber = buttonNumber + 1;
        let newObject = {
          'title': item.Title,
          'url': item.SliderImage.Url
        };
        objectArray.push(newObject);
      }
      buttons += `
      <button id="myButton_${buttonNumber}">${buttonNumber}</button>
      `
    });
    html += `
    <img class=${styles.imageNS} src="${objectArray[imageNumber].url}"/>
    <p class=${styles.titleNS}>${objectArray[imageNumber].title}</p>
    `
    const listContainer: Element = this.domElement.querySelector('#spListContainer');  
    listContainer.innerHTML = html;

    const listContainer2: Element = this.domElement.querySelector('#buttons');  
    listContainer2.innerHTML = buttons;

    for (let j = 1; j < objectArray.length +1; ++j) {
      var elem = document.getElementById('myButton_' + j);
      elem.addEventListener('click', function() {
        imageNumber = j-1;
          console.log('id: ' + j);
      });
    }
  };

  // private updateBtnNum(){
  //   console.log('check');
  // }
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
