import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IFirstWebpartProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  test: string;
  test1: boolean;
  test2: string;
  test3: boolean;
  context: WebPartContext
}
