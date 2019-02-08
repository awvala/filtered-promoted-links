import { IFilteredPromotedLinksWebPartProps } from '../IFilteredPromotedLinksWebPartProps';
import { SPHttpClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IFilteredPromotedLinksProps extends IFilteredPromotedLinksWebPartProps {
  isWorkbench: boolean;
  siteUrl: string;
  spHttpClient: SPHttpClient;
  //context: WebPartContext;
}

export interface IFilteredPromotedLinksProps {
  listName: string;
  filterName: string;
}

export interface IFilteredPromotedLinkDataItem {
  Title: string;
  ImageUrl: string;
  Description: string;
  LinkUrl: string;
  Owner: string;
}
