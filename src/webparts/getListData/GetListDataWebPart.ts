import { SPHttpClientResponse, SPHttpClient } from '@microsoft/sp-http';
import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown
} from '@microsoft/sp-webpart-base';

import { escape } from '@microsoft/sp-lodash-subset';

import styles from './GetListData.module.scss';
import MockHttpClient from './MockHttpClient';
import * as strings from 'getListDataStrings';

import { IGetListDataWebPartProps, ISPLists, ISPList } from './IGetListDataWebPartProps';

export default class GetListDataWebPart extends BaseClientSideWebPart<IGetListDataWebPartProps> {
  private sharepointLists: any[];
  public render(): void {
    this.domElement.innerHTML = `
      <div class="spListContainer" />
      </div>`;
    // Get the sharepoint lists to initialize the property pane
    if (!this.sharepointLists) {
      this._getSharepointLists()
        .then((res) => {
          this.sharepointLists = res.value;
          this.render();
        });
    } else {
      this._renderListAsync();
    }
  }

  private _renderListAsync(): void {
    // Local environment
    if (Environment.type === EnvironmentType.Local) {
      this._getMockData().then((res) => {
        this._renderList(res.value);
      });
    }
    else if (Environment.type == EnvironmentType.SharePoint ||
      Environment.type == EnvironmentType.ClassicSharePoint) {
      this._getListData()
        .then((response) => {
          this._renderList(response.value);
        });
    }
  }

  //Itmes from the list
  private _renderList(dataArr: ISPList[]): void {
    let html: string = ``;
    const listContainer: Element = this.domElement.querySelector('.spListContainer');
    try {
      dataArr.forEach((item: ISPList) => {
        html += `
      <ul class="${styles.list}">
          <li class="${styles.listItem}">
              <span class="ms-font-l">${item.Title}</span>
          </li>
      </ul>`;
      });
    }
    catch (error) {
      html += 'Unable to retreive data. error:' + error;
    }
    listContainer.innerHTML = html;
  }

  //Get data from sharepoint list
  private _getListData(): Promise<ISPLists> {
    const endpoint = this.context.pageContext.web.absoluteUrl +
      `/_api/web/lists/GetByTitle('${this.sharepointLists[this.properties.SharepointList] ? this.sharepointLists[this.properties.SharepointList].Title : 'testList'}')/items?$orderby=Id asc`;
    return this.context.spHttpClient
      .get(endpoint,
      SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  //Get Mock data 
  private _getMockData(): Promise<ISPLists> {
    return MockHttpClient.get()
      .then((data: ISPList[]) => {
        var listData: ISPLists = { value: data };
        return listData;
      }) as Promise<ISPLists>;
  }

  //Get all list from shorepoint context
  private _getSharepointLists(): Promise<ISPLists> {
    return this.context.spHttpClient
    .get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=Hidden eq false and BaseType ne 1`,
     SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  //Add lists to the Dropdown in the property panel
  private _transformListsIntoProperties(lists) {
    let listDropdown = [];
    if (lists) {
      lists.forEach((list, index) => {
        listDropdown.push(
          { key: index, text: list.Title }
        );
      });
    } else {
      listDropdown = [{ key: '1', text: 'Unable to retreive lists' }];
    }
    return listDropdown;
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const lists = this._transformListsIntoProperties(this.sharepointLists);

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
                PropertyPaneDropdown('SharepointList', {
                  label: 'Sharepoint List',
                  options: lists,
                  selectedKey: '1',
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
