import * as React from 'react';
import styles from './PnPControls.module.scss';
import { IPnPControlsProps } from './IPnPControlsProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class PnPControls extends React.Component<IPnPControlsProps, {}> {
  public render(): React.ReactElement<IPnPControlsProps> {
    return (
      <div className={ styles.pnPControls }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.title)}</p>
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
