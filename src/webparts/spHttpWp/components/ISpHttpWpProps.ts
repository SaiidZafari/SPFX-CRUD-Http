import { WebPartContext } from '@microsoft/sp-webpart-base';
// export interface ISpHttpWpProps {
//   description: string;
//   isDarkTheme: boolean;
//   environmentMessage: string;
//   hasTeamsContext: boolean;
//   userDisplayName: string;
// }

export interface ISpHttpWpProps {
  description: string;
  hasTeamsContext: boolean;
  context: WebPartContext;
}
