declare interface IGetListDataStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'getListDataStrings' {
  const strings: IGetListDataStrings;
  export = strings;
}
