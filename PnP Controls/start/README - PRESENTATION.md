# PnP Controls

This guide demonstrates how to use PnP reusable controls, property pane controls and PnPJS library on SPFx solutions.
It was created for demonstration purposes only. If creating a solution to be used on production, you need to also account for other scenarios (like using data services and mock data, tests, error handling, etc)

## Before the demo

Ensure you start from the "start" folder under the PnP Controls directory.

The start folder already contains a React web part with the required additional modules installed

* PnPJS
* PnP reusable controls
* PnP reusable property-pane controls

All you have to do is run "npm install" to install the required node modules.

### Steps if starting from an empty project

1. Install PnPJS

```TypeScript
npm install @pnp/logging @pnp/common @pnp/odata @pnp/sp @pnp/graph --save
```

Import library and setup on web part init (as per PnPJS documentation)

2. Install PnP reusable controls

```TypeScript
npm install @pnp/spfx-controls-react --save --save-exact
```

3. Install PnP reusable property-pane controls

```TypeScript
npm install @pnp/spfx-property-controls --save --save-exact
```

## Demo

Start by checking the imported references to property pane controls and PnPJS into your web part

```TypeScript
// PnP imports
import { sp } from "@pnp/sp";
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { PropertyFieldTermPicker, IPickerTerms } from '@pnp/spfx-property-controls/lib/PropertyFieldTermPicker';
```

Next, check that the web part properties include the options we need

* the web part title
* a reference to the source list
* a reference to the selected term

```TypeScript
export interface IPnPControlsWebPartProps {
  title: string;
  list: string;
  term: IPickerTerms;
}
```

Check that the React component properties include

* a reference to the web part context
* a reference to the web part display mode
* properties to handle the title update
* a reference to the source list
* a reference to the selected term

```TypeScript
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { DisplayMode } from '@microsoft/sp-core-library';
import { IPickerTerms } from "@pnp/spfx-property-controls/lib/PropertyFieldTermPicker";

export interface IPnPControlsProps {
  context: WebPartContext;
  displayMode: DisplayMode;
  title: string;
  updateTitle: (value: string) => void;
  list: string;
  term: IPickerTerms;
}
```

And finally, check that web part render function passes the correct properties to the React component

```TypeScript
context: this.context,
displayMode: this.displayMode,
title: this.properties.title,
updateTitle: (value: string) => {
  this.properties.title = value;
},
list: this.properties.list,
term: this.properties.term
```

Update property pane fields to include a list and term picker. Test the web part and validate that the property values are logged to the browser console

```TypeScript
PropertyFieldListPicker('list', {
  label: 'Select a list',
  selectedList: this.properties.list,
  includeHidden: false,
  orderBy: PropertyFieldListPickerOrderBy.Title,
  disabled: false,
  baseTemplate: 101, // filtering for document libraries
  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
  properties: this.properties,
  context: this.context,
  onGetErrorMessage: null,
  deferredValidationTime: 0,
  key: 'listPickerFieldId'
}),
PropertyFieldTermPicker('term', {
  label: 'Select a term',
  panelTitle: 'Select a term',
  initialValues: this.properties.term,
  allowMultipleSelections: false,
  excludeSystemGroup: false,
  limitByTermsetNameOrID: "Department",
  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
  properties: this.properties,
  context: this.context,
  onGetErrorMessage: null,
  deferredValidationTime: 0,
  key: 'termSetsPickerFieldId'
})
```

Check the required import statements in the React component

```TypeScript
import { sp } from "@pnp/sp";
import { Placeholder } from '@pnp/spfx-controls-react/lib/Placeholder';
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { ListView, IViewField } from '@pnp/spfx-controls-react/lib/ListView';
```

Check the interface with an items property to be used as the state interface
For simplicity, it can be created next to the properties interface

```TypeScript
export interface IPnPControlsState {
  items?: any[];
  noItemsPlaceholder?: boolean;
}
```

Check that the new interface is imported and added to our class declaration

```TypeScript
import { IPnPControlsProps, IPnPControlsState } from './IPnPControlsProps';

export default class PnPControls extends React.Component<IPnPControlsProps, IPnPControlsState> {
```

Check that the class constructor is used to set the initial state

```TypeScript
constructor(props: IPnPControlsProps) {
  super(props);

  this.state = {
    items: [],
    noItemsPlaceholder: false
  };
}
```

Update render method to return the Placeholder if no list is selected
(Add the code below as the first block of the render function and keep the existing code below)

```TypeScript
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
```

Implement _configureWebPart function to open the web part property pane and test the changes

```TypeScript
private _onConfigure() {
  this.props.context.propertyPane.open();
}
```

Next, add a function to retrieve list items using PnPJS
The function should also filter items by the selected term

```TypeScript
private async _getItems() {
  let select = '*';
  let expand = 'File';
  let filter = '';

  // filter by selected term if required
  if (this.props.term !== undefined && this.props.term !== null && this.props.term.length > 0) {
    const term = this.props.term[0];

    select = `${select},TaxCatchAll/Term`;
    expand = `${expand},TaxCatchAll`;
    filter = `TaxCatchAll/Term eq '${term.name}'`;
  }

  const items = await sp.web.lists.getById(this.props.list).items
    .select(select)
    .expand(expand)
    .filter(filter)
    .get();

  // update state
  this.setState({
    items: items ? items : [],
    noItemsPlaceholder: items.length === 0
  });
  console.log('List Items:', this.state.items);
}
```

Call the _getItems function during the React component lifecycle and test the changes

```TypeScript
public componentDidMount() {
  if (this.props.list !== null && this.props.list !== "" && this.props.list !== undefined) {
    this._getItems();
  }
}

public componentDidUpdate(prevProps: IPnPControlsProps, prevState: IPnPControlsState) {
  if (this.props.list !== prevProps.list || this.props.term !== prevProps.term) {
    if (this.props.list !== null && this.props.list !== "" && this.props.list !== undefined) {
      this._getItems();
    }
  }
}
```

Add the WebPartTitle component to the main block of the render method (replace default content generated by default)

```TypeScript
<WebPartTitle
  displayMode={this.props.displayMode}
  title={this.props.title}
  updateProperty={this.props.updateTitle} />
```

Next we are going to display the items as a list.
Create an object that uses the IViewField[] interface. This object will define the columns to be rendered.

```TypeScript
private _viewFields: IViewField[] = [
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
```

Add the ListView to the render method. We are also dynamically displaying a different Placeholder for when no items are available.

Replace the render function as below

```TypeScript
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
    this.state.noItemsPlaceholder ?
      (
        <Placeholder
          iconName="InfoSolid"
          iconText="No items found"
          description="No items to display" />
      ) : (
        <div>
          <WebPartTitle displayMode={this.props.displayMode}
            title={this.props.title}
            updateProperty={this.props.updateTitle} />
          <ListView items={this.state.items}
            viewFields={this._viewFields}
            selection={this._getSelection}
            iconFieldName="File.ServerRelativeUrl" />
        </div>
      )
  );
}
```

Finally, add a new function to log the List items selected by the user

```TypeScript
private _getSelection(items: any[]) {
  console.log('Selected List items:', items);
}
```
