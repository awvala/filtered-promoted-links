import * as React from 'react';
import styles from './FilteredPromotedLinks.module.scss';
import { IFilteredPromotedLinksProps, IFilteredPromotedLinkDataItem } from './IFilteredPromotedLinksProps';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/components/Spinner';
import { Placeholder } from '@pnp/spfx-controls-react/lib/Placeholder';
import FilteredPromotedLinkItem, { IFilteredPromotedLinkItemProps } from './FilteredPromotedLinksItem';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient } from '@microsoft/sp-http';

export interface IFilteredPromotedLinksState {
  listData: IFilteredPromotedLinkDataItem[];
  loading?: boolean;
  showPlaceholder?: boolean;
}

export default class FilteredPromotedLinks extends React.Component<IFilteredPromotedLinksProps, IFilteredPromotedLinksState> {

  constructor(props: IFilteredPromotedLinksProps, state: IFilteredPromotedLinksState) {
    super(props);

    this._onConfigure = this._onConfigure.bind(this);

    this.state = {
      listData: [],
      loading: false,
      showPlaceholder: (this.props.listName === null || this.props.listName === ""),
    };
  }

  /*
   * Opens the web part property pane
  */
  private _onConfigure() {
    this.props.context.propertyPane.open();
  }

  public render(): React.ReactElement<IFilteredPromotedLinksProps> {

    if (this.state.showPlaceholder) {
      // Check if placeholder needs to be shown
      return (
        <Placeholder
          iconName="Edit"
          iconText="Filtered Promoted Links web part configuration"
          description="Please configure the web part to show filtered promoted links."
          buttonLabel="Configure"
          onConfigure={this._onConfigure}
        />
      );
    }

    return (
      <div className={styles.filteredPromotedLinks}>
        <div >
          <h2>{this.props.description}</h2>
        </div>
        {
          this.props.missingField ?
            (
              <Placeholder
                iconName="InfoSolid"
                iconText="The list you selected is missing a Filter or Owner field."
                description="Please select another list or add the missing field(s) to this list."
                buttonLabel="Configure"
                onConfigure={this._onConfigure}
              />
            ) : this.state.loading ?
              (
                <Spinner size={SpinnerSize.large} label="Retrieving results ..." />
              ) : this.state.listData.length === 0 ?
                (
                  <Placeholder
                    iconName="InfoSolid"
                    iconText="No items found"
                    description="Please select a new list or update the filter in the property pane."
                    buttonLabel="Configure"
                    onConfigure={this._onConfigure}
                  />
                ) :
                (
                  <div className={styles.container}>
                    {
                      this.state.listData.map((item: IFilteredPromotedLinkDataItem) => {
                        return <FilteredPromotedLinkItem
                          title={item.Title}
                          description={item.Description}
                          imageUrl={item.ImageUrl}
                          href={item.LinkUrl}
                          owner={item.Owner}
                        />;
                      })
                    }
                    <div style={{ clear: 'both' }}></div>
                  </div>
                )
        }
      </div>
    );
  }

  public componentDidMount(): void {
    if (this.props.listName !== null && this.props.listName !== "") {
      this.loadData();
    }
  }

  public componentDidUpdate(prevProps: IFilteredPromotedLinksProps, prevState: IFilteredPromotedLinksState, prevContext: any) {
    if (prevProps.listName != this.props.listName || prevProps.filterName != this.props.filterName) {
      if (this.props.listName !== null && this.props.listName !== "") {
        this.loadData();
      } else {
        this.setState({
          showPlaceholder: true
        });
      }
    }
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
    } else if (!this.props.missingField && this.props.filterName !== undefined) {
      // get data from SharePoint
      this.props.spHttpClient.get(`${this.props.siteUrl}/_api/Web/Lists(guid'${this.props.listName}')/items?$select=Title,Description,BackgroundImageLocation,LinkLocation,Owner/Title&$expand=Owner/Id&$filter=Category eq '${this.props.filterName}'`, SPHttpClient.configurations.v1)
        .then(response => {
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
              Owner: items.value[i].Owner.Title
            });
          }
          this.setState({
            listData: listItems,
            loading: false,
            showPlaceholder: false
          });
        }, (err: any) => {
          console.log(err);
        });
    } else {
      //  disable the Filter dropdown
      this.setState({
        listData: [],
        loading: false,
        showPlaceholder: false
      });
    }
  }
}