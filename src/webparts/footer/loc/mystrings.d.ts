declare interface IFooterWebPartStrings {
  PropertyPaneDescription: string;
  BrandGroupName: string;
  LinksGroupName: string;
  BrandFieldLabel: string;
  BlurbFieldLabel: string;
  CopyrightFieldLabel: string;
  ListFieldLabel: string;
  SocialFieldLabel: string;
  FieldHint: string;
}

declare module 'FooterWebPartStrings' {
  const strings: IFooterWebPartStrings;
  export = strings;
}
