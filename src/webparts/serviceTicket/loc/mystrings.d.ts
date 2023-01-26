declare interface IServiceTicketWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'ServiceTicketWebPartStrings' {
  const strings: IServiceTicketWebPartStrings;
  export = strings;
}
