declare interface ISearchExtensionLibraryStrings {
  Layouts: {
    EbscoList: {
      ShowFileIconLabel: string;
    },
    ShellReportsList: {
      ShowFileIconLabel: string;
    }
  }
}

declare module 'SearchExtensionLibraryStrings' {
  const strings: ISearchExtensionLibraryStrings;
  export = strings;
}
