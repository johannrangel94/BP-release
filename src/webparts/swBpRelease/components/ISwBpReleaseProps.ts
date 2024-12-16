import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ISwBpReleaseProps {
  context: WebPartContext;
  description: string;
  selectedList: string;
}
