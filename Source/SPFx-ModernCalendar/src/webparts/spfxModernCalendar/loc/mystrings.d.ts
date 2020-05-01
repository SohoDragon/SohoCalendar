declare interface ISpfxModernCalendarWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  WebPartTitleLabel: string;
  EventBackgroundColor: string;
  EventTitleColor: string;
  ListInfoGroupName: string;
  ListTitleLabel: string;
  StartDateFieldLabel: string;
  EndDateFieldLabel: string;
  EventTitleFieldLabel: string;
  EventDescriptionFieldLabel: string;
  AllDaysEventFieldLabel: string;
}

declare module 'SpfxModernCalendarWebPartStrings' {
  const strings: ISpfxModernCalendarWebPartStrings;
  export = strings;
}
