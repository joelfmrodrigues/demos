import * as React from 'react';
import styles from './HelloWorld.module.scss';
import { IHelloWorldProps } from './IHelloWorldProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { ListView, IViewField, SelectionMode, GroupOrder, IGrouping } from "@pnp/spfx-controls-react/lib/ListView";
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

const viewFields: IViewField[] = [
  {
    name: "Id",
    displayName: "ID",
    maxWidth: 25,
    minWidth: 25,
    sorting: true
  },
  {
    name: "File.Name",
    linkPropertyName: "File.ServerRelativeUrl",
    displayName: "Name",
    sorting: true
  }
];

export interface IHelloWorldState {
  items: any;
}

export default class HelloWorld extends React.Component<IHelloWorldProps, IHelloWorldState> {

constructor(props: IHelloWorldProps) {
  super(props);

  this.state = {
    items: []
  };
}


  public async componentDidMount(): Promise<void> {
    await this._getItems();
  }

  public render(): React.ReactElement<IHelloWorldProps> {
    return this.props.lists && this.props.lists.length > 0 ?
      (
        <ListView
          items={this.state.items}
          viewFields={viewFields}
          iconFieldName="ServerRelativeUrl"
          compact={true}
          selectionMode={SelectionMode.multiple}
          selection={this._getSelection}
          showFilter={true}
          defaultFilter=""
          filterPlaceHolder="Search..." />
      ) : (
        <Placeholder iconName='Edit'
          iconText='Configure your web part'
          description='Please configure the web part.'
          buttonLabel='Configure'
          onConfigure={this._onConfigure} />
      );
  }

  private _getSelection(items: any[]) {
    console.log('Selected items:', items);
  }

  private _onConfigure() {
    // Context of the web part
    this.props.context.propertyPane.open();
  }

  private _getItems = async () => {
    // get all the items from a list
    const items: any[] = await sp.web.lists.getById(this.props.lists).items.expand('File').select('ID,File').get();
    console.log(items);

    this.setState({
      items: items
    });
  }
}
