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
