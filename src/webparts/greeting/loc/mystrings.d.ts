declare interface IGreetingWebPartStrings {
  PropertyPaneDescription: string;
  AppearanceGroupName: string;
  ContentGroupName: string;
  LayoutFieldLabel: string;
  ChipsFieldLabel: string;
  FieldHint: string;
}

declare module 'GreetingWebPartStrings' {
  const strings: IGreetingWebPartStrings;
  export = strings;
}
