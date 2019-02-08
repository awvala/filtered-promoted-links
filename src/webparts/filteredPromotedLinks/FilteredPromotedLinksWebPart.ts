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
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';

export interface IFilteredPromotedLinksWebPartProps {
  description: string;
}

export default class FilteredPromotedLinksWebPart extends BaseClientSideWebPart<IFilteredPromotedLinksWebPartProps> {

  private lists: IPropertyPaneDropdownOption[];
  private listsDropdownDisabled: boolean = true;
  private filters: IPropertyPaneDropdownOption[];
  private filtersDropdownDisabled: boolean = true;

  public render(): void {
    const element: React.ReactElement<IFilteredPromotedLinksProps> = React.createElement(
      FilteredPromotedLinks,
      {
        isWorkbench: Environment.type == EnvironmentType.Local,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        spHttpClient: this.context.spHttpClient,
        listName: this.properties.listName,
        filterName: this.properties.filterName,
        description: this.properties.description,
        context: this.context
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneDropdown('listName', {
                  label: strings.ListNameFieldLabel,
                  options: this.lists,
                  disabled: this.listsDropdownDisabled
                }),
                PropertyPaneDropdown('filterName', {
                  label: strings.FilterNameFieldLabel,
                  options: this.filters,
                  disabled: this.filtersDropdownDisabled
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
    this.filtersDropdownDisabled = !this.properties.listName || !this.filters;

    if (this.lists) {
      return;
    }

    // Disable 02/07/19 as render method returned a blank.
    // this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'lists');

    this.fetchOptions()
      .then((data: IPropertyPaneDropdownOption[]): Promise<IPropertyPaneDropdownOption[]> => {
        this.lists = data;
        this.listsDropdownDisabled = false;
        this.context.propertyPane.refresh();
        return this.fetchFilterOptions();
      })
      .then((filterOptions: IPropertyPaneDropdownOption[]): void => {
        this.filters = filterOptions;
        this.filtersDropdownDisabled = !this.properties.listName;
        this.context.propertyPane.refresh();
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);
        this.render();
      });
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'listName' &&
      newValue) {
      // push new list value
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      // get previously selected filter
      const previousfilter: string = this.properties.filterName;
      // reset selected filter
      this.properties.filterName = undefined;
      // push new filter value
      this.onPropertyPaneFieldChanged('filterName', previousfilter, this.properties.filterName);
      // disable filter selector until new filters are loaded
      this.filtersDropdownDisabled = true;
      // refresh the filter selector control by repainting the property pane
      this.context.propertyPane.refresh();
      // communicate loading filters // Disable 02/07/19 as render method returned a blank.
      // this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'filters');

      this.fetchFilterOptions()
        .then((filterOptions: IPropertyPaneDropdownOption[]): void => {
          // store filters
          this.filters = filterOptions;
          // enable filter selector
          this.filtersDropdownDisabled = false;
          // clear Filter property pane field
          this.properties.filterName = "";
          // clear status indicator
          this.context.statusRenderer.clearLoadingIndicator(this.domElement);
          // re-render the web part as clearing the loading indicator removes the web part body
          this.render();
          // refresh the item selector control by repainting the property pane
          this.context.propertyPane.refresh();
        });
    }
    else {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    }
  }

  private fetchLists(url: string): Promise<ISPLists> {
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
    const url = this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=BaseTemplate eq 170 and Hidden eq false`;

    return this.fetchLists(url).then((response) => {
      let options: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();
      let lists: ISPList[] = response.value;
      lists.forEach((list: ISPList) => {
        // console.log("Found list with title = " + list.Title);
        options.push({ key: list.Id, text: list.Title });
      });

      return options;
    });
  }

  //  SharePoint API
  private fetchFilterOptions(): Promise<IPropertyPaneDropdownOption[]> {
    const url = this.context.pageContext.web.absoluteUrl + `/_api/Web/Lists(guid'${this.properties.listName}')/items?$select=Filter&$orderby=Filter asc`;
    
    if (!this.properties.listName) {
      // resolve to empty options since no list has been selected
      return Promise.resolve();
    } else {
      return this.fetchLists(url).then((response) => {
        let options: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();
        let lists: ISPList[] = response.value;
        lists.forEach((list: ISPList) => {
          options.push({ key: list.Filter, text: list.Filter });
        });
        // Remove duplicate filters
        options = options.filter((value, index, array) => 
          !array.filter((v, i) => JSON.stringify(value) == JSON.stringify(v) && i < index).length);
        return options;
      });
    }
  }

}

/*
// Static method
private loadfilters(): Promise<IPropertyPaneDropdownOption[]> {
  console.log ("start Loadfilters");
  if (!this.properties.listName) {
    // resolve to empty options since no list has been selected
    return Promise.resolve();
  }

  const wp: FilteredPromotedLinksWebPart = this;

  return new Promise<IPropertyPaneDropdownOption[]>((resolve: (options: IPropertyPaneDropdownOption[]) => void, reject: (error: any) => void) => {
    setTimeout(() => {
      const filters = {
        sharedDocuments: [
          {
            key: 'spfx_presentation.pptx',
            text: 'SPFx for the masses'
          },
          {
            key: 'hello-world.spapp',
            text: 'hello-world.spapp'
          }
        ],
        myDocuments: [
          {
            key: 'isaiah_cv.docx',
            text: 'Isaiah CV'
          },
          {
            key: 'isaiah_expenses.xlsx',
            text: 'Isaiah Expenses'
          }
        ]
      };
      resolve(filters[wp.properties.listName]);
    }, 2000);
    // console.log("In Loadfilters method: "+ this.filters);
  });
}*/
