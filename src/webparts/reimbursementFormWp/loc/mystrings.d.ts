declare interface IReimbursementFormWpWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;

  AdminApprover : string;
  CategoryList : string;
  ClientList : string;
  DepartmentsList : string;
  ProgramList : string;
  ProjectList : string;
  ReimbursementRequestList : string;
  ReimbursementItemsList : string;
  ReimbursementRequestSettingsList : string;
  SubcategoryList : string;
  TasksList : string;
}

declare module 'ReimbursementFormWpWebPartStrings' {
  const strings: IReimbursementFormWpWebPartStrings;
  export = strings;
}
