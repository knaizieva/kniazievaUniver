declare interface IPlannerReportsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  LibraryFieldLabel: string;
  PlanFieldLabel: string;
  GroupFieldLabel: string;
  CreateReportButton: string;
  Loading: string;
  SegmentLabel: string;
  TasksLabel: string;
  CompletedTaskLabel: string;
  SetWebPartPropsMessage: string;
  SetPropertiesLabel: string;
  FileCreatedMessage: string;
  FileLabel:string;
  FileNotCreatedMessage: string;
}

declare module 'PlannerReportsWebPartStrings' {
  const strings: IPlannerReportsWebPartStrings;
  export = strings;
}
