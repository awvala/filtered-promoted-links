import { IFilteredPromotedLinksWebPartProps } from '../IFilteredPromotedLinksWebPartProps';
import { SPHttpClient } from '@microsoft/sp-http';

export interface IFilteredPromotedLinksProps {
  listName: string;
  categoryName: string;
}

export interface IFilteredPromotedLinksProps extends IFilteredPromotedLinksWebPartProps {
  isWorkbench: boolean;
  siteUrl: string;
  spHttpClient: SPHttpClient;
}

export interface IFilteredPromotedLinkDataItem {
  Title: string;
  ImageUrl: string;
  Description: string;
  LinkUrl: string;
  Owner: string;
}
