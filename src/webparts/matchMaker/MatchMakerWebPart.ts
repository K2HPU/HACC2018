//Needed to import ‘Environment’ and the ‘EnvironmentType’ modules to implement get main repository
import {  
  Environment,  
  EnvironmentType ,
  Version
} from '@microsoft/sp-core-library';

import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './MatchMakerWebPart.module.scss';
import * as strings from 'MatchMakerWebPartStrings';

//Needed to import the ‘MockHttpClient’ module
import MockHttpClient from './MockHttpClient'; 

//Needed to import spHttpClient to call REST API requests 
import {  
  SPHttpClient,
  SPHttpClientResponse,
  MSGraphClient
} from '@microsoft/sp-http'; 

// These are the types for graph nodes that are published separetlely (User field types, messages, contacts, etc.)
// To reference Microsoft Graph types, see directions at https://github.com/microsoftgraph/msgraph-typescript-typings/
// The dependency has been added in package.json, so just run npm install
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";

export interface IMatchMakerWebPartProps {
  description: string;
}
//Added list interface to retrieve items form Main Repository
export interface ISPLists {  
  value: ISPList[];  
}  
export interface ISPList {  
  Title: string;  
  Resource_x0020_Type: string;  
  Subject_x0020_Area: string;  
  Target_x0020_Audience: string;  
}    

export default class MatchMakerWebPart extends BaseClientSideWebPart<IMatchMakerWebPartProps> {

  //Added the mock list item retrieval method 
  private _getMockListData(): Promise<ISPLists> {  
    return MockHttpClient.get(this.context.pageContext.web.absoluteUrl).then(() => {  
        const listData: ISPLists = {  
            value:  
            [  
                { Title: 'FileName1', Resource_x0020_Type: 'Lesson Plan', Subject_x0020_Area: 'Math', Target_x0020_Audience: '1stGrade' },  
                { Title: 'FileName2', Resource_x0020_Type: 'Instructions', Subject_x0020_Area: 'English', Target_x0020_Audience: '2ndGrade' },  
                { Title: 'FileName3', Resource_x0020_Type: 'Field Trip', Subject_x0020_Area: 'Science', Target_x0020_Audience: '3rdGrade'  }  
            ]  
            };  
        return listData;  
    }) as Promise<ISPLists>;  
} 

// private _getUserData(): Promise<IUser> {  
//   return this.context.spHttpClient.get(`https://ksbetest.sharepoint.com/sites/dev/ociss/_api/web/lists/GetByTitle('Main Repository')/Items`, SPHttpClient.configurations.v1)  
//       .then((response: SPHttpClientResponse) => {   
//         debugger;  
//         return response.json();  
//       });  
//     } 

//Added this method to get SharePoint list items, using REST API
  private _getListData(): Promise<ISPLists> {  
    return this.context.spHttpClient.get(`https://ksbetest.sharepoint.com/sites/dev/ociss/_api/web/lists/GetByTitle('Main Repository')/Items`, SPHttpClient.configurations.v1)  
        .then((response: SPHttpClientResponse) => {   
          debugger;  
          return response.json();  
        });  
    } 
//Used to check Environment.type value and if it is equal to Environment.Local, the MockHttpClient method, which returns dummy data which will be called else the method that calls REST API is able to retrieve SharePoint list items will be called
    private _renderListAsync(): void {  
      
      if (Environment.type === EnvironmentType.Local) {  
        this._getMockListData().then((response) => {  
          this._renderList(response.value);  
        });  
      }  
       else {  
         this._getListData()  
        .then((response) => {  
          this._renderList(response.value);  
        });  
     }  
  } 

  //add this method to create HTML table out of the retrieved SharePoint list items
  private _renderList(items: ISPList[]): void {

    let html: string = '<table class="TFtable" border=1 width=100% style="border-collapse: collapse;">';  
    var i = 0;
    html += `<tr>`;  
    items.forEach((item: ISPList) => {  
      if (i < 3) {
        html += `  
        <td width="33%"><ul>  
            <li>${item.Title}</li>  
            <li>${item.Resource_x0020_Type}</li>  
            <li>${item.Subject_x0020_Area}</li>  
            <li>${item.Target_x0020_Audience}</li>  
            </ul></td> 
            `;
      }
      i++;
    });  
    html += `</tr></table>`;  
    const listContainer: Element = this.domElement.querySelector('#spListContainer');  
    listContainer.innerHTML = html;  
  } 

  
  //Replace Render method to enable rendering of the list items  
  public render(): void {  
    // this.context.msGraphClientFactory
    // .getClient()
    // .then((client: MSGraphClient): void => {
    //   // get information about the current user from the Microsoft Graph
    //   client
    //     // .api('/me')
    //     // .get((error, response: any, rawResponse?: any) => {

    //     // handle the response
    //     .api('/me')
    //     .select("displayName")
    //     .get((err, user:MicrosoftGraph.User) => {
    //         if (err) {
    //             console.error(err);
    //             return;
    //         }            
    //         user.displayName;
    //     });          
    // });


    this.domElement.innerHTML = `  
    <div class="${styles.matchMaker}">  
    <div class="${styles.container}">  
      <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">  
        <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">  
          <span class="ms-font-xl ms-fontColor-white" style="font-size:28px">G5 MatchMaker App</span>  
            
          <p class="ms-font-l ms-fontColor-white" style="text-align: center">Demo : Retrieve Main Repository Data from SharePoint List</p>  
        </div>  
      </div>  
      <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">  
      <div style="background-color:Black;color:white;text-align: center;font-weight: bold;font-size:18px;">Resources</div>  
      <br>  
        <div id="spListContainer" />  
      </div>  
    </div>  
    </div>`;  
    this._renderListAsync();
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
