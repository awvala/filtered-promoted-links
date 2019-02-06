import * as React from 'react';
import styles from './FilteredPromotedLinks.module.scss';
import { IFilteredPromotedLinksProps, IFilteredPromotedLinkDataItem } from './IFilteredPromotedLinksProps';
import FilteredPromotedLinkItem, { IFilteredPromotedLinkItemProps } from './FilteredPromotedLinksItem';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient } from '@microsoft/sp-http';

export interface IFilteredPromotedLinksState {
  listData: IFilteredPromotedLinkDataItem[];
}

export default class FilteredPromotedLinks extends React.Component<IFilteredPromotedLinksProps, IFilteredPromotedLinksState> {

  constructor(props: IFilteredPromotedLinksProps, state: IFilteredPromotedLinksState) {
    super(props);

    this.state = { listData: [] };
  }

  public render(): React.ReactElement<IFilteredPromotedLinksProps> {
    return (
      <div className={styles.filteredPromotedLinks}>
        <div className={styles.container}>

          {
            this.state.listData.map((item: IFilteredPromotedLinkDataItem) => {
              return <FilteredPromotedLinkItem
                title={item.Title}
                description={item.Description}
                imageUrl={item.ImageUrl}
                href={item.LinkUrl}
                owner={item.Owner} />;

            })
          }

          <div style={{ clear: 'both' }}></div>
        </div>
      </div>
    );
  }

  public componentDidMount(): void {
    this.loadData();
  }

  private loadData(): void {
    if (this.props.isWorkbench) {
      // get mock data in Workbench
      this.setState({
        listData: [
          {
            Title: "Test Item",
            Description: "Test description",
            ImageUrl: "https://media-cdn.tripadvisor.com/media/photo-s/04/a8/17/f5/el-arco.jpg",
            LinkUrl: "http://www.google.com",
            Owner: "Jim Davis"
          },
          {
            Title: "Test Item with a Long Title",
            Description: "Test description",
            ImageUrl: "https://pgcpsmess.files.wordpress.com/2014/04/330277-red-fox-kelly-lyon-760x506.jpg",
            LinkUrl: "http://www.google.com",
            Owner: "Jim Davis"
          },
          {
            Title: "Test Item",
            Description: "Test description",
            ImageUrl: "https://s-media-cache-ak0.pinimg.com/736x/d6/d4/d7/d6d4d7224687ca3de4a160f5264b5b99.jpg",
            LinkUrl: "Test item with a long description for display.",
            Owner: "Jim Davis"
          }
        ]
      });
    } else {
      // get data from SharePoint
      this.props.spHttpClient.get(`${this.props.siteUrl}/_api/Web/Lists(guid'${this.props.listName}')/items?$filter=Category eq '${this.props.categoryName}'`, SPHttpClient.configurations.v1)
        .then(response => {
          // console.log(`Get SharePoint list data: ${response}`);
          return response.json();
        })
        .then((items: any) => {
          // console.log(items);
          const listItems: IFilteredPromotedLinkDataItem[] = [];
          for (let i: number = 0; i < items.value.length; i++) {
            listItems.push({
              Title: items.value[i].Title,
              Description: items.value[i].Description,
              ImageUrl: items.value[i].BackgroundImageLocation.Url,
              LinkUrl: items.value[i].LinkLocation.Url,
              Owner: items.value[i].OwnerId
            });
          }
          this.setState({ listData: listItems });
        }, (err: any) => {
          console.log(err);
        });
    }
  }

  public componentDidUpdate(prevProps: IFilteredPromotedLinksProps, prevState: IFilteredPromotedLinksState, prevContext: any) {
    if (prevProps.listName != this.props.listName) {
      this.loadData();
    }
  }
}