declare interface IClSearchStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'clSearchStrings' {
  const strings: IClSearchStrings;
  export = strings;
}
