import { WebPartContext } from "@microsoft/sp-webpart-base";
import{SPHttpClient} from "@microsoft/sp-http";

export interface IFileViewerProps {
  listTitle: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context:WebPartContext;
  siteUrl:string;
  spHttpClient: SPHttpClient;
  close:() => void;
}
