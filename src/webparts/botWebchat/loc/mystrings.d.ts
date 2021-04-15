declare interface IBotWebPartStrings {
  BasicGroupName: string;
  ConnectionGroupName: string;
  PropertyPaneBotButtonText: string; 
  BotButtonTextFieldLabel: string;
  PropertyPaneChatWindowHeaderTitle: string; 
  ChatWindowHeaderTitleFieldLabel: string;
  PropertyPaneDescription: string; 
  DescriptionFieldLabel: string;
  PropertyPaneConnectBy: string;
  ConnectByFieldLabel: string;
  PropertyPaneDLSecret: string;
  DLSecretFieldLabel: string;
}

declare module 'BotWebPartStrings' {
  const strings: IBotWebPartStrings;
  export = strings;
}
