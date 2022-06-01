declare interface IBirthdayWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  NumberUpComingDaysLabel: string;
}

declare module "BirthdayWebPartStrings" {
  const strings: IBirthdayWebPartStrings;
  export = strings;
}
