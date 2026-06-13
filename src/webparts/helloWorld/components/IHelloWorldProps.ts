export interface IProject {
  id: string;
  name: string;
  description?: string;
  colour: string;
  documentCount?: number;
  lastModified?: string;
}

export interface IHelloWorldProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
