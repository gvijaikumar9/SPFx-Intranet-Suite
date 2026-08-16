declare interface IChartFromListWebPartStrings {
  PropertyPaneDescription: string;
  AppearanceGroupName: string;
  DataGroupName: string;
  TitleFieldLabel: string;
  LayoutFieldLabel: string;
  MaxBarsLabel: string;
  UseDemoLabel: string;
  ListFieldLabel: string;
  CategoryFieldLabel: string;
  ValueFieldLabel: string;
  FieldHint: string;
}

declare module 'ChartFromListWebPartStrings' {
  const strings: IChartFromListWebPartStrings;
  export = strings;
}
