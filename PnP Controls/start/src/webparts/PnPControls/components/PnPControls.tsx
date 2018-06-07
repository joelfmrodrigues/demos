import * as React from 'react';
import styles from './PnPControls.module.scss';
import { IPnPControlsProps, IPnPControlsState } from './IPnPControlsProps';
import { escape } from '@microsoft/sp-lodash-subset';

// PnP imports
import { sp } from "@pnp/sp";
import { Placeholder } from '@pnp/spfx-controls-react/lib/Placeholder';
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { ListView, IViewField } from '@pnp/spfx-controls-react/lib/ListView';

export default class PnPControls extends React.Component<IPnPControlsProps, IPnPControlsState> {

  // change 6


  constructor(props: IPnPControlsProps) {
    super(props);

    this.state = {
      items: [],
    };
  }

  // change 7
  public render(): React.ReactElement<IPnPControlsProps> {

    console.log('List Items:', this.state.items);

    // change 2


    return (
      // change 5
      <div className={styles.pnPControls}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Welcome to SharePoint!</span>
              <p className={styles.subTitle}>Customize SharePoint experiences using Web Parts.</p>
              <p className={styles.description}>{escape(this.props.title)}</p>
              <a href="https://aka.ms/spfx" className={styles.button}>
                <span className={styles.label}>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }

  // change 3


  // change 4


  // change 8

}
