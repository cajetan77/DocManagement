import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IDocumentLibraryViewsProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  siteUrl: string;
  listTitle: string;
  context: WebPartContext;
}
