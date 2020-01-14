import * as React from 'react';
import styles from './RndWebpart.module.scss';
import { IRndWebpartProps } from './IRndWebpartProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class RndWebpart extends React.Component<IRndWebpartProps, {}> {
  public render(): React.ReactElement<IRndWebpartProps> {
    return (
      <div className={ styles.rndWebpart }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>

              <div dangerouslySetInnerHTML={{ __html:this.props.htmlCode }} />          

            </div>
          </div>
        </div>
      </div>
    );
  }
}
