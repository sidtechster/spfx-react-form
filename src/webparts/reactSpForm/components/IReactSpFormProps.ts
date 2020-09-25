import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IReactSpFormProps {
  listName: string;
  context: WebPartContext;
  siteUrl: string;
}
