import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import { sp, SPRest } from "@pnp/sp";
import { PropertyFieldTermPicker, ICheckedTerms } from '@pnp/spfx-property-controls/lib/PropertyFieldTermPicker';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';

import * as strings from 'PnPControlsWebPartStrings';
import PnPControls from './components/PnPControls';
import { IPnPControlsProps } from './components/IPnPControlsProps';

export interface IPnPControlsWebPartProps {
  sp: SPRest;
  listId: string;
  term: ICheckedTerms;
}

export default class PnPControlsWebPart extends BaseClientSideWebPart<IPnPControlsWebPartProps> {

  public onInit(): Promise<void> {

    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      });
    });
  }

  public render(): void {

    console.info('List Id:', this.properties.listId);
    console.info('Term:', this.properties.term);

    const element: React.ReactElement<IPnPControlsProps> = React.createElement(
      PnPControls,
      {
        context: this.context,
        sp: sp,
        listId: this.properties.listId,
        term: this.properties.term
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
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
              ]
            }
          ]
        }
      ]
    };
  }
}
