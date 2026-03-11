import { SPHttpClient } from '@microsoft/sp-http';

export interface IEntitySearchWebpartProps {
  description: string;
  listId: string;
  primaryFieldInternalName: string;
  secondaryFieldInternalName: string;
  tertiaryFieldInternalName: string;
  badgeFieldInternalName: string;
  actionsConfigurationJson: string;
  spHttpClient: SPHttpClient;
  siteUrl: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
