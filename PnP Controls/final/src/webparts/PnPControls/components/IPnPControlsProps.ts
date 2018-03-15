import { WebPartContext } from "@microsoft/sp-webpart-base";
import { DisplayMode } from '@microsoft/sp-core-library';
import { SPRest } from "@pnp/sp";
import { ICheckedTerms } from "@pnp/spfx-property-controls/lib/PropertyFieldTermPicker";

export interface IPnPControlsProps {
  context: WebPartContext;
  displayMode: DisplayMode;
  sp: SPRest;
  title: string;
  updateTitle: (value: string) => void;
  list: string;
  term: ICheckedTerms;
}

export interface IPnPControlsState {
  items?: any[];
}
