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
  private views: IPropertyPaneDropdownOption[];
  private viewsDropdownDisabled: boolean = true;

  public render(): void {
    const element: React.ReactElement<IFilteredPromotedLinksProps> = React.createElement(
      FilteredPromotedLinks,
      {
        listName: this.properties.listName,
        viewName: this.properties.viewName
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
                PropertyPaneDropdown('viewName', {
                  label: strings.ViewNameFieldLabel,
                  options: this.views,
                  disabled: this.viewsDropdownDisabled
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
    this.viewsDropdownDisabled = !this.properties.listName || !this.views;

    if (this.lists) {
      return;
    }

    this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'lists');

    this.fetchOptions()
      .then((data: IPropertyPaneDropdownOption[]): Promise<IPropertyPaneDropdownOption[]> => {
        this.lists = data;
        this.listsDropdownDisabled = false;
        this.context.propertyPane.refresh();
        return this.fetchViewOptions();
      })
      .then((viewOptions: IPropertyPaneDropdownOption[]): void => {
        this.views = viewOptions;
        this.viewsDropdownDisabled = !this.properties.listName;
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
      // get previously selected view
      const previousView: string = this.properties.viewName;
      // reset selected view
      this.properties.viewName = undefined;
      // push new view value
      this.onPropertyPaneFieldChanged('viewName', previousView, this.properties.viewName);
      // disable view selector until new views are loaded
      this.viewsDropdownDisabled = true;
      // refresh the view selector control by repainting the property pane
      this.context.propertyPane.refresh();
      // communicate loading views
      this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'views');

      this.fetchViewOptions()
        .then((viewOptions: IPropertyPaneDropdownOption[]): void => {
          // store views
          this.views = viewOptions;
          // enable view selector
          this.viewsDropdownDisabled = false;
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
  private fetchViewOptions(): Promise<IPropertyPaneDropdownOption[]> {
    const url = this.context.pageContext.web.absoluteUrl + `/_api/Web/Lists(guid'${this.properties.listName}')/Views`;

    if (!this.properties.listName) {
      // resolve to empty options since no list has been selected
      return Promise.resolve();
    } else {
    // console.log(`Start fetchViews: ${url}`);
    return this.fetchLists(url).then((response) => {
      let options: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();
      let lists: ISPList[] = response.value;
      lists.forEach((list: ISPList) => {
        // console.log("Found list with title = " + list.Title);
        options.push({ key: list.Id, text: list.Title });
      });

      return options;
    })
  }
  }

}

/*
// Static method
private loadViews(): Promise<IPropertyPaneDropdownOption[]> {
  console.log ("start LoadViews");
  if (!this.properties.listName) {
    // resolve to empty options since no list has been selected
    return Promise.resolve();
  }

  const wp: FilteredPromotedLinksWebPart = this;

  return new Promise<IPropertyPaneDropdownOption[]>((resolve: (options: IPropertyPaneDropdownOption[]) => void, reject: (error: any) => void) => {
    setTimeout(() => {
      const views = {
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
      resolve(views[wp.properties.listName]);
    }, 2000);
    // console.log("In LoadViews method: "+ this.views);
  });
}*/
