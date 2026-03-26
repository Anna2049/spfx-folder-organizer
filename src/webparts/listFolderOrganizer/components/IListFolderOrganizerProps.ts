import { SPHttpClient } from "@microsoft/sp-http";

export interface IListFolderOrganizerProps {
  spHttpClient: SPHttpClient;
  siteUrl: string;
}
