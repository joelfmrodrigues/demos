import * as React from 'react';
import styles from './PnPControls.module.scss';
import { IPnPControlsProps, IPnPControlsState } from './IPnPControlsProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { Placeholder } from '@pnp/spfx-controls-react/lib/Placeholder';

export default class PnPControls extends React.Component<IPnPControlsProps, IPnPControlsState> {

  constructor(props: IPnPControlsProps) {
    super(props);

    this.state = {
      items: [],
    };
  }

  public componentDidMount() {
    if (this.props.list !== null && this.props.list !== "" && this.props.list === undefined) {
      this._getItems();
    }
  }

  public componentDidUpdate(prevProps: IPnPControlsProps, prevState: IPnPControlsState) {
    if (this.props.list !== prevProps.list || this.props.term !== prevProps.term) {
      if (this.props.list !== null && this.props.list !== "" && this.props.list === undefined) {
        this._getItems();
      }
    }
  }

  public render(): React.ReactElement<IPnPControlsProps> {
    if (this.props.list === null || this.props.list === "" || this.props.list === undefined) {
      return (
        <Placeholder
          iconName="Edit"
          iconText="Configure your web part"
          description="Please configure the web part."
          buttonLabel="Configure"
          onConfigure={this._onConfigure.bind(this)} />
      );
    }
    return (
      <div className={styles.pnPControls}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Welcome to SharePoint!</span>
              <p className={styles.subTitle}>Customize SharePoint experiences using Web Parts.</p>
              <a href="https://aka.ms/spfx" className={styles.button}>
                <span className={styles.label}>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }

  private _onConfigure() {
    this.props.context.propertyPane.open();
  }

  private async _getItems() {
    let select = '*';
    let expand = 'File';
    let filter = '';

    // filter by selected term if required
    if (this.props.term !== undefined && this.props.term !== null) {
      const term = this.props.term[0];

      select = `${select},TaxCatchAll/Term`;
      expand = `${expand},TaxCatchAll`;
      filter = `TaxCatchAll/Term eq '${term.name}'`;
    }

    const items = await this.props.sp.web.lists.getById(this.props.list).items
      .select(select)
      .expand(expand)
      .filter(filter)
      .get();

    console.log('List Items:', items);

    // update state
    this.setState({
      items: items ? items : []
    });
  }
}
