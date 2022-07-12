import { IReadonlyTheme } from "@microsoft/sp-component-base";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IReadReceiptWebpartProps {
  documentTitle: string;
  currentUserDisplayName: string;
  storageList: string;
  acknowledgementLabel: string;
  acknowledgementMessage: string;
  readMessage: string;
  themeVariant: IReadonlyTheme | undefined;
  configured: boolean;
  context: WebPartContext;
  // description: string;
  // isDarkTheme: boolean;
  // environmentMessage: string;
  // hasTeamsContext: boolean;
  // userDisplayName: string;
}
