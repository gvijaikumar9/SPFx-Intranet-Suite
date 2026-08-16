declare interface IKudosWebPartStrings {
  PropertyPaneDescription: string;
  AppearanceGroupName: string;
  DataGroupName: string;
  TitleFieldLabel: string;
  LayoutFieldLabel: string;
  ShowPhotoLabel: string;
  MaxItemsLabel: string;
  UseDemoLabel: string;
  ListFieldLabel: string;
  FieldHint: string;
}

declare module 'KudosWebPartStrings' {
  const strings: IKudosWebPartStrings;
  export = strings;
}
