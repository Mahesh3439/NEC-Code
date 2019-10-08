import * as React from 'react';
import styles from './PromotionResponse.module.scss';
import { IPromotionResponseProps } from './IPromotionResponseProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class PromotionResponse extends React.Component<IPromotionResponseProps, {}> {
  public render(): React.ReactElement<IPromotionResponseProps> {
    return (
      <div className={ styles.promotionResponse }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
             
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
