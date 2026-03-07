import { SPHttpClient } from '@microsoft/sp-http';

export interface ICommandCalendarProps {
  description: string;
  calendarSources: string;
  swimLaneMappings: string;
  categoryColorMappings: string;
  lookBackDays: number;
  lookAheadDays: number;
  spHttpClient: SPHttpClient;
}
