
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ISpFXCrudProps {
  description: string;
  hasTeamsContext: boolean;
  context: WebPartContext;
}
