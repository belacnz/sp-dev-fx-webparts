declare interface IFileUploadWebPartStrings {
  ErrorOnLoadingWebDropdown: string;
  LoadingWebDropdown: string;
  WebUrlFieldLabel: string;
    ErrorWebNotFound: string;
    ErrorWebAccessDenied: string;
    WebUrlFieldPlaceholder: string;
    SiteUrlFieldPlaceholder: string;
  ErrorOnLoadingSiteDropdown: string;
  LoadingSiteDropdown: string;
  SiteUrlFieldLabel: string;
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'FileUploadWebPartStrings' {
  const strings: IFileUploadWebPartStrings;
  export = strings;
}
