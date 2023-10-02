declare interface IExportListItemsCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'ExportListItemsCommandSetStrings' {
  const strings: IExportListItemsCommandSetStrings;
  export = strings;
}
