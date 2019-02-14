import { SPHttpClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IFilteredPromotedLinksWebPartProps {
    listName: string;  // Stores the Promoted Links list name
    filterName: string;  // Stores the filter field value
    description: string; // Stores the Web Part Title
    context: WebPartContext;
    missingField: boolean; // Stores state of Filter API call
   }

   export interface ISPList {
    Title: string;
    Id: string;
    Category: string;
  }
  
  export interface ISPLists {
    value: ISPList[];
  }
