import { ServiceScope } from "@microsoft/sp-core-library";
import { HttpClient } from "@microsoft/sp-http";

export interface ICreateOfferSettingsProps {
  serviceScope: ServiceScope;
  httpClient: HttpClient;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  mySiteUrl: string;
  userLogin: string;
}
