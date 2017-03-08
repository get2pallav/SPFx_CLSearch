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
import {IListService} from './Services/MockResultSourceService';
import {MockListService} from './Services/MockResultSourceService';
import {ListService} from './Services/MockResultSourceService';

export default class ClSearchWebPart extends BaseClientSideWebPart<IClSearchWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.helloWorld}">
        <div class="${styles.container}">
          <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
            <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <span class="ms-font-xl ms-fontColor-white">Welcome to SharePoint!</span>
              <p class="ms-font-l ms-fontColor-white">Customize SharePoint experiences using Web Parts.</p>
              <p class="ms-font-l ms-fontColor-white">${escape(this.properties.description)}</p>
              <a href="https://aka.ms/spfx" class="${styles.button}">
                <span class="${styles.label}">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;
     (new ListService()).getLists();
  }
 
  public onInit<T>():Promise<T> {
    debugger;
    let mockService = Environment.type == EnvironmentType.SharePoint ?  new ListService()  : new MockListService();
    this._options = [];
    return new Promise<T>((resolve) => {
      mockService.getLists().then((lists) => {
        lists.forEach((list) => {
            this._options.push(<IPropertyPaneDropdownOption>{
              text:list,
            });
          resolve(undefined);
        });
      });
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
