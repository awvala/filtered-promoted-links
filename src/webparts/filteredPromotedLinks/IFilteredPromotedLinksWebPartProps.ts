import { SPHttpClient } from '@microsoft/sp-http';

export interface IFilteredPromotedLinksWebPartProps {
    listName: string;  // Stores the Promoted Links list name
    filterName: string;  // Stores the filter field value
    description: string; // Stores the Web Part Title
   }

   export interface ISPList {
    Title: string;
    Id: string;
    Filter: string;
  }
  
  export interface ISPLists {
    value: ISPList[];
  }
