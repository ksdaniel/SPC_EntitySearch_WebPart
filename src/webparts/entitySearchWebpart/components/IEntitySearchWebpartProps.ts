import { SPHttpClient } from '@microsoft/sp-http';

export interface IEntitySearchWebpartProps {
  description: string;
  listId: string;
  titleFieldInternalName: string;
  typeFieldInternalName: string;
  dealFieldInternalName: string;
  statusFieldInternalName: string;
  spHttpClient: SPHttpClient;
  siteUrl: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
