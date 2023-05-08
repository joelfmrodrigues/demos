import { BaseComponentContext } from '@microsoft/sp-component-base';

export interface IHelloWorldProps {
  spfxContext: BaseComponentContext;
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}

export interface ISite {
  Title: string;
}

export interface IHelloWorldState {
  items: ISite[];
}