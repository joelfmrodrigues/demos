import { SPRest } from "@pnp/sp";
import { ICheckedTerms } from "@pnp/spfx-property-controls/lib/PropertyFieldTermPicker";

export interface IPnPControlsProps {
  sp: SPRest;
  listId: string;
  term: ICheckedTerms;
}
