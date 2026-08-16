declare interface IWeatherWebPartStrings {
  PropertyPaneDescription: string;
  AppearanceGroupName: string;
  ContentGroupName: string;
  TitleFieldLabel: string;
  LayoutFieldLabel: string;
  LocationFieldLabel: string;
  UnitsFieldLabel: string;
  FieldHint: string;
}

declare module 'WeatherWebPartStrings' {
  const strings: IWeatherWebPartStrings;
  export = strings;
}
