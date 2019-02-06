import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-webpart-base';

import * as strings from 'FilteredPromotedLinksWebPartStrings';
import FilteredPromotedLinks from './components/FilteredPromotedLinks';
import { IFilteredPromotedLinksProps } from './components/IFilteredPromotedLinksProps';
import { IFilteredPromotedLinksWebPartProps, ISPLists, ISPList } from './IFilteredPromotedLinksWebPartProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IFilteredPromotedLinksWebPartProps {
  description: string;
}

export default class FilteredPromotedLinksWebPart extends BaseClientSideWebPart<IFilteredPromotedLinksWebPartProps> {

  private lists: IPropertyPaneDropdownOption[];
  private listsDropdownDisabled: boolean = true;

  public render(): void {
    const element: React.ReactElement<IFilteredPromotedLinksProps > = React.createElement(
      FilteredPromotedLinks,
      {
        listName: this.properties.listName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
                PropertyPaneDropdown('listName', {
                  label: strings.ListNameFieldLabel,
                  options: this.lists,
                  disabled: this.listsDropdownDisabled
                })
              ]
            }
          ]
        }
      ]
    };
  }

  protected onPropertyPaneConfigurationStart(): void {
    this.listsDropdownDisabled = !this.lists;

    if (this.lists) {
      return;
    }

    this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'lists');

    this.fetchOptions()
    .then((data: IPropertyPaneDropdownOption[]) /*: Promise<IPropertyPaneDropdownOption[]>*/ => {
      this.lists = data;
      this.listsDropdownDisabled = false;
      this.context.propertyPane.refresh();
      //return this.loadViews();
    });
  }

  private fetchLists(url: string) : Promise<ISPLists> {
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
        if (response.ok) {
          return response.json();
        } else {
          console.log("WARNING - failed to hit URL " + url + ". Error = " + response.statusText);
          return null;
        }
      });
  }

  private fetchOptions(): Promise<IPropertyPaneDropdownOption[]> {
    var url = this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=BaseTemplate eq 170 and Hidden eq false`;

    return this.fetchLists(url).then((response) => {
        var options: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();
        var lists: ISPList[] = response.value;
        lists.forEach((list: ISPList) => {
            console.log("Found list with title = " + list.Title);
            options.push( { key: list.Id, text: list.Title });
        });

        return options;
    });
  }
}
