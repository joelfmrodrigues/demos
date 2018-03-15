import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPRest } from "@pnp/sp";
import { ICheckedTerms } from "@pnp/spfx-property-controls/lib/PropertyFieldTermPicker";

export interface IPnPControlsProps {
  context: WebPartContext;
  sp: SPRest;
  list: string;
  term: ICheckedTerms;
}

export interface IPnPControlsState {
  items?: any[];
}
