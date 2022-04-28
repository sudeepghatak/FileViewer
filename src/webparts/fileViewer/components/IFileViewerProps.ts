import { WebPartContext } from "@microsoft/sp-webpart-base";
import{SPHttpClient} from "@microsoft/sp-http";

export interface IFileViewerProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context:WebPartContext;
  siteUrl:string;
  spHttpClient: SPHttpClient;
}
