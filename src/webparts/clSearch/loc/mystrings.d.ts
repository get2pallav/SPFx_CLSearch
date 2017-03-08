declare interface IClSearchStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  SearchResultSources:string;
}

declare module 'clSearchStrings' {
  const strings: IClSearchStrings;
  export = strings;
}
