# PnP Controls

This guide demonstrates how to use PnP reusable controls, property pane controls and PnPJS library on SPFx solutions.
It was created for demonstration purposes only. If creating a solution to be used on production, you need to also account for other scenarios (like using data and mock services, tests, error handling, etc).

## Before the demo

Ensure you start from the "start" folder under the PnP Controls directory.

The start folder already contains a React web part with the required additional modules installed

* PnPJS
* PnP reusable controls
* PnP reusable property-pane controls

### Steps if starting from an empty project

1. Install PnPJS

npm install @pnp/logging @pnp/common @pnp/odata @pnp/sp @pnp/graph --save

Import library and setup on web part init (as per PnPJS documentation)

2. Install PnP reusable controls

npm install @pnp/spfx-controls-react --save --save-exact

Configure resource file by adding the below to config/config.json
"ControlStrings": "node_modules/@pnp/spfx-controls-react/lib/loc/{locale}.js"

3. Install PnP reusable property-pane controls

npm install @pnp/spfx-property-controls --save --save-exact

Configure resource file by adding the below to config/config.json
"PropertyControlStrings": "node_modules/@pnp/spfx-property-controls/lib/loc/{locale}.js"

## Demo

Update web part properties to include

* a configured reference to PnPJS,
* a reference to the source list
* a reference to the selected term

```TypeScript
export interface IPnPControlsWebPartProps {
  sp: SPRest;
  listId: string;
  term: ICheckedTerms;
}
```

Update web part imports to resolve missing references and include property pane controls

```TypeScript
import { sp, SPRest } from "@pnp/sp";
import { ICheckedTerms } from "@pnp/spfx-property-controls/lib/PropertyFieldTermPicker";
import { PropertyFieldTermPicker, ICheckedTerms } from '@pnp/spfx-property-controls/lib/PropertyFieldTermPicker';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
```

Update property pane fields to include a list and term picker

```TypeScript
PropertyFieldListPicker('listId', {
  label: 'Select a list',
  selectedList: this.properties.listId,
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
  onPropertyChange: this.onPropertyPaneFieldChanged,
  properties: this.properties,
  context: this.context,
  onGetErrorMessage: null,
  deferredValidationTime: 0,
  key: 'termSetsPickerFieldId'
})
```

Add logging to the render function and validate that the properties are being populated

```TypeScript
console.info('List Id:', this.properties.listId);
console.info('Term:', this.properties.term);
```

Update React component properties to include

* a configured reference to PnPJS,
* a reference to the source list
* a reference to the selected term

```TypeScript
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPRest } from "@pnp/sp";
import { ICheckedTerms } from "@pnp/spfx-property-controls/lib/PropertyFieldTermPicker";

export interface IPnPControlsProps {
  context: WebPartContext;
  sp: SPRest;
  listId: string;
  term: ICheckedTerms;
}
```

Update web part render function to pass the correct properties to the React component

```TypeScript
context: this.context,
sp: sp,
listId: this.properties.listId,
term: this.properties.term
```

Import Placeholder control into the React component

```TypeScript
import { Placeholder } from '@pnp/spfx-controls-react/lib/Placeholder';
```

Update render method to return the Placeholder if no list is selected
(Add the code below as the first block of the render function and keep the existing code below)

```TypeScript
if (this.props.listId === null || this.props.listId === "" || this.props.listId === undefined) {
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








