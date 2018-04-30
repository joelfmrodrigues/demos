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

Configure resource file by adding the below to config/config.json

```TypeScript
"ControlStrings": "node_modules/@pnp/spfx-controls-react/lib/loc/{locale}.js"
```

3. Install PnP reusable property-pane controls

```TypeScript
npm install @pnp/spfx-property-controls --save --save-exact
```

Configure resource file by adding the below to config/config.json

```TypeScript
"PropertyControlStrings": "node_modules/@pnp/spfx-property-controls/lib/loc/{locale}.js"
```

## Demo

Start by importing references to property pane controls and PnPJS into your web part

```TypeScript
import { sp, SPRest } from "@pnp/sp";
import { ICheckedTerms } from "@pnp/spfx-property-controls/lib/PropertyFieldTermPicker";
import { PropertyFieldTermPicker, IPickerTerms } from '@pnp/spfx-property-controls/lib/PropertyFieldTermPicker';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
```

Next, update web part properties to include the options we need

* a configured reference to PnPJS,
* the web part title
* a reference to the source list
* a reference to the selected term

```TypeScript
export interface IPnPControlsWebPartProps {
  sp: SPRest;
  title: string;
  list: string;
  term: IPickerTerms;
}
```

Update property pane fields to include a list and term picker

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
  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
  properties: this.properties,
  context: this.context,
  onGetErrorMessage: null,
  deferredValidationTime: 0,
  key: 'termSetsPickerFieldId'
})
```

Add logging to the render function and validate that the properties are being populated

```TypeScript
console.info('List Id:', this.properties.list);
console.info('Term:', this.properties.term);
```

Update React component properties to include

* a reference to the web part context
* a reference to the web part display mode
* a configured reference to PnPJS,
* properties to handle the title update
* a reference to the source list
* a reference to the selected term

```TypeScript
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { DisplayMode } from '@microsoft/sp-core-library';
import { SPRest } from "@pnp/sp";
import { IPickerTerms } from "@pnp/spfx-property-controls/lib/PropertyFieldTermPicker";

export interface IPnPControlsProps {
  context: WebPartContext;
  displayMode: DisplayMode;
  sp: SPRest;
  title: string;
  updateTitle: (value: string) => void;
  list: string;
  term: IPickerTerms;
}
```

Update web part render function to pass the correct properties to the React component

```TypeScript
context: this.context,
displayMode: this.displayMode,
sp: sp,
title: this.properties.title,
updateTitle: (value: string) => {
  this.properties.title = value;
},
list: this.properties.list,
term: this.properties.term
```

Import Placeholder control into the React component

```TypeScript
import { Placeholder } from '@pnp/spfx-controls-react/lib/Placeholder';
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

Implement _configureWebPart function to open the web part property pane

```TypeScript
private _onConfigure() {
  this.props.context.propertyPane.open();
}
```

Create a new interface with an items property to be used as the state interface
For simplicity, it can be created next to the properties interface

```TypeScript
export interface IPnPControlsState {
  items?: any[];
}
```

Import the new interface and add it to our class declaration

```TypeScript
import { IPnPControlsProps, IPnPControlsState } from './IPnPControlsProps';

export default class PnPControls extends React.Component<IPnPControlsProps, IPnPControlsState> {
```

Add a constructor to the class and set the initial state

```TypeScript
constructor(props: IPnPControlsProps) {
  super(props);

  this.state = {
    items: [],
  };
}
```

Next, add a function to query list items using PnPJS
The function should also filter items by the selected term

```TypeScript
private async _getItems() {
  let select = '*';
  let expand = 'File';
  let filter = '';

  console.log('Selected Term: ', this.props.term);

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
```

Call the _getItems function during the React component lifecycle

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

Import the WebPartTitle component

```TypeScript
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
```

Add it to the main block of the render method (replace default content generated by default)

```TypeScript
<WebPartTitle displayMode={this.props.displayMode}
  title={this.props.title}
  updateProperty={this.props.updateTitle} />
```

Next we are going to display the items as a list

Start by importing the required components

```TypeScript
import { ListView, IViewField } from '@pnp/spfx-controls-react/lib/ListView';
```

Next, create an object that uses the IViewField[] interface. This object will define the columns to be rendered.

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

Add a new function to log the List items selected by the user

```TypeScript
private _getSelection(items: any[]) {
  console.log('Selected List items:', items);
}
```

Finally, add the ListView to the render method. We are also dynamically displaying a different Placeholder for when no items are available.

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
    this.state.items.length === 0 ?
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