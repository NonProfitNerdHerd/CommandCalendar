declare interface ICommandCalendarWebPartStrings {
  PropertyPaneDescription: string;
  GeneralSettingsGroupName: string;
  CalendarSourcesGroupName: string;
  DescriptionFieldLabel: string;
  LookBackDaysFieldLabel: string;
  LookAheadDaysFieldLabel: string;
  CalendarSourcesFieldLabel: string;
  CalendarSourcesFieldDescription: string;
  CategoryColorMappingsFieldLabel: string;
  CategoryColorMappingsFieldDescription: string;
}

declare module 'CommandCalendarWebPartStrings' {
  const strings: ICommandCalendarWebPartStrings;
  export = strings;
}
