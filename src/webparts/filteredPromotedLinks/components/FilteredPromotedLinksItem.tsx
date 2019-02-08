import * as React from 'react';
import styles from './FilteredPromotedLinks.module.scss';
import { Image, IImageProps, ImageFit } from 'office-ui-fabric-react';
import { escape } from '@microsoft/sp-lodash-subset';

export interface IFilteredPromotedLinkItemProps {
  imageUrl: string;
  title: string;
  description: string;
  href: string;
  owner: string;
}

export interface IFilteredPromotedLinkItemState {
    hovering: boolean;
  }

  export default class FilteredPromotedLinks extends React.Component<IFilteredPromotedLinkItemProps, IFilteredPromotedLinkItemState> {

    constructor(props: IFilteredPromotedLinkItemProps, state: IFilteredPromotedLinkItemState) {
      super(props);
  
      this.state = {
        hovering: false
      };
    }
  
    public mouseOver(event): void {
      this.setState({ hovering: true });
    }
  
    public mouseOut(event): void {
      this.setState({ hovering: false });
    }
  
    public render(): React.ReactElement<IFilteredPromotedLinkItemProps> {
      return (
        <a href={this.props.href} target="_top" role="listitem"
              onMouseOver={this.mouseOver.bind(this)} onMouseOut={this.mouseOut.bind(this)}>
          <div className={styles.pLinkItemWrapper}>
            <Image className={styles.pLinkItemImage} src={this.props.imageUrl} shouldFadeIn={true} imageFit={ImageFit.cover} />
            <div className={this.state.hovering ? styles.pLinkItemHoverPanelExpanded : styles.pLinkItemHoverPanel}>
              <div className={styles.pLinkItemTitle}>{this.props.title}</div>
              <div className={styles.pLinkItemDesc}>
                <p className={styles.pLinkItemOwner}>{this.props.owner}</p>
                {this.props.description}
              </div>
            </div>
          </div>
        </a>
      );
    }
  }