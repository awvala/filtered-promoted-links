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
import { IFilteredPromotedLinksWebPartProps } from './IFilteredPromotedLinksWebPartProps';

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
}
