import { SPHttpClient } from '@microsoft/sp-http';

export interface IFilteredPromotedLinksWebPartProps {
    listName: string;
    categoryName: string;
   }

   export interface ISPList {
    Title: string;
    Id: string;
    Category: string;
  }
  
  export interface ISPLists {
    value: ISPList[];
  }
