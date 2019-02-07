import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
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
  private categories: IPropertyPaneDropdownOption[];
  private categoriesDropdownDisabled: boolean = true;

  public render(): void {
    const element: React.ReactElement<IFilteredPromotedLinksProps> = React.createElement(
      FilteredPromotedLinks,
      {
        isWorkbench: Environment.type == EnvironmentType.Local,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        spHttpClient: this.context.spHttpClient,
        listName: this.properties.listName,
        categoryName: this.properties.categoryName
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
                }),
                PropertyPaneDropdown('categoryName', {
                  label: strings.CategoryNameFieldLabel,
                  options: this.categories,
                  disabled: this.categoriesDropdownDisabled
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
    this.categoriesDropdownDisabled = !this.properties.listName || !this.categories;

    if (this.lists) {
      return;
    }

    // this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'lists');

    this.fetchOptions()
      .then((data: IPropertyPaneDropdownOption[]): Promise<IPropertyPaneDropdownOption[]> => {
        this.lists = data;
        this.listsDropdownDisabled = false;
        this.context.propertyPane.refresh();
        return this.fetchCategoryOptions();
      })
      .then((categoryOptions: IPropertyPaneDropdownOption[]): void => {
        this.categories = categoryOptions;
        this.categoriesDropdownDisabled = !this.properties.listName;
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
      // get previously selected category
      const previouscategory: string = this.properties.categoryName;
      // reset selected category
      this.properties.categoryName = undefined;
      // push new category value
      this.onPropertyPaneFieldChanged('categoryName', previouscategory, this.properties.categoryName);
      // disable category selector until new categories are loaded
      this.categoriesDropdownDisabled = true;
      // refresh the category selector control by repainting the property pane
      this.context.propertyPane.refresh();
      // communicate loading categories
      // this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'categories');

      this.fetchCategoryOptions()
        .then((categoryOptions: IPropertyPaneDropdownOption[]): void => {
          // store categories
          this.categories = categoryOptions;
          // enable category selector
          this.categoriesDropdownDisabled = false;
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
  private fetchCategoryOptions(): Promise<IPropertyPaneDropdownOption[]> {
    const url = this.context.pageContext.web.absoluteUrl + `/_api/Web/Lists(guid'${this.properties.listName}')/items?$select=Category&$orderby=Category asc`;
    
    if (!this.properties.listName) {
      // resolve to empty options since no list has been selected
      return Promise.resolve();
    } else {
      // console.log(`Start fetchcategories: ${url}`);
      return this.fetchLists(url).then((response) => {
        //console.log(`category response: ${response}`);
        let options: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();
        let lists: ISPList[] = response.value;
        // console.log(`Lists: ${lists}`);
        lists.forEach((list: ISPList) => {
          options.push({ key: list.Category, text: list.Category });
        });
        // Remove duplicate Categories
        options = options.filter((value, index, array) => 
          !array.filter((v, i) => JSON.stringify(value) == JSON.stringify(v) && i < index).length);
        // console.log(`Options: ${options}`);
        return options;
      });
    }
  }

}

/*
// Static method
private loadcategories(): Promise<IPropertyPaneDropdownOption[]> {
  console.log ("start Loadcategories");
  if (!this.properties.listName) {
    // resolve to empty options since no list has been selected
    return Promise.resolve();
  }

  const wp: FilteredPromotedLinksWebPart = this;

  return new Promise<IPropertyPaneDropdownOption[]>((resolve: (options: IPropertyPaneDropdownOption[]) => void, reject: (error: any) => void) => {
    setTimeout(() => {
      const categories = {
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
      resolve(categories[wp.properties.listName]);
    }, 2000);
    // console.log("In Loadcategories method: "+ this.categories);
  });
}*/
