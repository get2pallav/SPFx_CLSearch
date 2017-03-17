import { Version,
Environment,
EnvironmentType } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './ClSearch.module.scss';
import * as strings from 'clSearchStrings';
import { IClSearchWebPartProps } from './IClSearchWebPartProps';
import * as SearchService from './Services/SearchService';
 import {SPComponentLoader} from "@microsoft/sp-loader";
 

export default class ClSearchWebPart extends BaseClientSideWebPart<IClSearchWebPartProps> {

public constructor(){
  super();
  SPComponentLoader.loadCss("https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.2.0/css/fabric.min.css");
    SPComponentLoader.loadCss("https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.2.0/css/fabric.components.min.css");
}

  public render(): void {

    this.domElement.innerHTML = `
      <div class="${styles.helloWorld}">
        <div class="${styles.container}">
          <div class="ms-Grid-row ms-bgColor-themeLight ms-fontColor-white ${styles.row}">
            <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <div>
                <span class="ms-font-xl ms-fontColor-white">${escape(this.properties.description)}</span><br/>
               <div class="ms-Grid"> 
                    <div class="ms-Grid-row">
                      <div class="ms-Grid-col ms-u-sm10"><input class="ms-TextField-field" id="textInput" placeholder="Search..." /></div>
                      <div class="ms-Grid-col ms-u-sm2"> <Button class="ms-Button ms-Button--primary" id="btnSearchSubmit" type="submit" value="Submit">Search</Button></div>
                    </div>
                  </div>
                 <div class="ms-List ms-Grid-col ms-u-sm12" id="searchResults"></div>
                </div>
            </div>
          </div>         
        </div>
      </div>`;
     this.AttachEvents();

  }

  private AttachEvents():void{
   const btnSearch = this.domElement.querySelector("#btnSearchSubmit");

   const queryText:HTMLElement = <HTMLInputElement>this.domElement.querySelector("#textInput");
   btnSearch.addEventListener('click',() => {
        (new ClSearchWebPart()).handleOnChange(queryText);
   });
  }
 
 public handleOnChange(text:HTMLElement):void{
  (new ClSearchWebPart()).renderResults((<HTMLInputElement>text).value)
     .then((html) =>{
       const element  = document.getElementById("searchResults");
       element.innerHTML = html;
     });
 } 

 private renderResults(query:string):Promise<string>{
   const _search:SearchService.ISearchService = Environment.type == EnvironmentType.SharePoint ?
                                                   new SearchService.SearchService() : new SearchService.MockSearchService();
   let resultsHtml:string = '';

   return new Promise<string>((resolve) => {
     if(query){
        _search.GetSearchResults(query,"", 3, 0)
        .then((results) => {
            results.forEach((result) => {
                resultsHtml += `<div class=""ms-ListItem ms-Grid-col ms-u-sm8">
                                 <a href="${result.link}"><span class="ms-ListItem-primaryText" >${result.title}</span></a>
                                  <span class="ms-ListItem-secondaryText">${result.author}<span>
                                 <span class="ms-ListItem-tertiaryText">${result.description}</span>
                                 <span class="ms-ListItem-metaText">10:15a</span>
                                  <div class="ms-ListItem-actions">
                                      <div class="ms-ListItem-action" targerUrl="${result.link}"><i class="ms-Icon ms-Icon--OpenInNewWindow">
                                      </i></div>
                                    </div>
                               </div>`;
                  });
              })
       .then(
         () => {
          setTimeout(() => {
            const action:HTMLCollectionOf<Element> = document.getElementsByClassName("ms-ListItem-action");
            for(let i=0;i<action.length;i++){
              action[i].addEventListener('click',(e)=>{
                window.open((e.currentTarget as Element).getAttribute("targerUrl"));
              });
            }
          },300);
           resolve(resultsHtml);
          }
       );
     }
     else{
       resultsHtml += "Please provide Search query..";
       resolve(resultsHtml);
     }
   });
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

