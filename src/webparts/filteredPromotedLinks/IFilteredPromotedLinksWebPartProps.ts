export interface IFilteredPromotedLinksWebPartProps {
    listName: string;
    viewName: string;
   }

   export interface ISPList {
    Title: string;
    Id: string;
  }
  
  export interface ISPLists {
    value: ISPList[];
  }