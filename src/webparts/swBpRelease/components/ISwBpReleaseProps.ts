import { WebPartContext } from "@microsoft/sp-webpart-base";
//import { IFieldInfo } from "@pnp/sp/fields";

export interface ISwBpReleaseProps {
  context: WebPartContext;
  selectedList: string;
  selectedFields: string[];
  orderedItems: Array<any>;
}
