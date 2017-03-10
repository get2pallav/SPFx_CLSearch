import { Version,
Environment,
EnvironmentType } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './ClSearch.module.scss';
import * as strings from 'clSearchStrings';
import { IClSearchWebPartProps } from './IClSearchWebPartProps';
import * as CLListService from './Services/MockResultSourceService';
import * as SearchService from './Services/SearchService';

export default class ClSearchWebPart extends BaseClientSideWebPart<IClSearchWebPartProps> {

  public render(): void {

    this.domElement.innerHTML = `
      <div class="${styles.helloWorld}">
        <div class="${styles.container}">
          <div class="ms-Grid-row ms-bgColor-themeTertiary ms-fontColor-white ${styles.row}">
            <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <span class="ms-font-xl ms-fontColor-white">Welcome to SharePoint!</span>
              <p class="ms-font-l ms-fontColor-white">${escape(this.properties.description)}</p>
              <p class="ms-font-l ms-fontColor-white">${escape(this.properties.resultsource)}</p>
              <div id="searchResults"></div>
            </div>
          </div>
        </div>
      </div>`;
     this.renderResults()
     .then((html) =>{
       let element  =  this.domElement.querySelector("#searchResults");
       element.innerHTML = html;
     });
  }
 
 private renderResults():Promise<string>{
   const _search:SearchService.ISearchService = Environment.type == EnvironmentType.SharePoint ?
                                                   new SearchService.SearchService() : new SearchService.MockSearchService();
   let resultsHtml:string = '';

   return new Promise<string>((resolve) => {

        _search.GetSearchResults('','','',10)
        .then((results) => {
            results.forEach((result) => {
                resultsHtml += `<div class="cob-result">
                                  <div class="">
                                      <i class="ms-Icon ms-Icon--PageSolid"></i>
                                      <a href="${result.link}"><span class="${styles.label}" >${result.title}</span></a>
                                      
                                  </div>
                                  <p class="ms-fontSize-s">${result.description}</p>
                              </div>`;
                  });
              })
       .then(
         () => {resolve(resultsHtml);}
       );
   });
 }
 
  protected onPropertyPaneConfigurationStart():void{
  let mockService = Environment.type == EnvironmentType.SharePoint ?  new CLListService.ListService()  : new CLListService.MockListService();
  this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'Result Sources');

  this._options = [];
     mockService.getLists().then((lists) => {
        lists.forEach((list) => {
            this._options.push(<IPropertyPaneDropdownOption>{
              text:list,
             key:list
            });
        });
      this.context.propertyPane.refresh();
      this.context.statusRenderer.clearLoadingIndicator(this.domElement);
     });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private _options:IPropertyPaneDropdownOption[];

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
                }),
                PropertyPaneDropdown('resultsource',{
                  label:strings.SearchResultSources,
                  options:this._options
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
