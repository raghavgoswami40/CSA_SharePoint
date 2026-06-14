import { SPHttpClient } from '@microsoft/sp-http';

export interface IProject {
  id: string;
  name: string;
  siteUrl: string;
  description?: string;
  colour: string;
  lastModified?: string;
}

export interface ISiteSearchResult {
  title: string;
  url: string;
}

export interface IHelloWorldProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  spHttpClient: SPHttpClient;
  webUrl: string;
}
