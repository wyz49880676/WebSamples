declare interface IRecentDocsWebPartStrings {
  PropertyPaneDescription: string;
  
  InfoGroupName: string;
  WPIconURLLabel: string;
  WPTitleLabel: string;

  DataGroupName: string;
  SiteURLLabel: string;
  ListURLLabel: string;
  ListNameLabel: string;
}

declare module 'RecentDocsWebPartStrings' {
  const strings: IRecentDocsWebPartStrings;
  export = strings;
}
