import { ServiceScope } from "@microsoft/sp-core-library";

export interface IOfferCreationSpFxProps {
  siteUrl: string;
  siteDomain: string;
  serviceScope: ServiceScope;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
