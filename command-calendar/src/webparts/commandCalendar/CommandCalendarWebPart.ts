import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneSlider,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'CommandCalendarWebPartStrings';
import CommandCalendar from './components/CommandCalendar';
import { ICommandCalendarProps } from './components/ICommandCalendarProps';

export interface ICommandCalendarWebPartProps {
  description: string;
  calendarSources: string;
  swimLaneMappings: string;
  categoryColorMappings: string;
  lookBackDays: number;
  lookAheadDays: number;
}

export default class CommandCalendarWebPart extends BaseClientSideWebPart<ICommandCalendarWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ICommandCalendarProps> = React.createElement(
      CommandCalendar,
      {
        description: this.properties.description,
        calendarSources: this.properties.calendarSources,
        swimLaneMappings: this.properties.swimLaneMappings,
        categoryColorMappings: this.properties.categoryColorMappings,
        lookBackDays: this.properties.lookBackDays,
        lookAheadDays: this.properties.lookAheadDays,
        spHttpClient: this.context.spHttpClient
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    if (!this.properties.lookBackDays && this.properties.lookBackDays !== 0) {
      this.properties.lookBackDays = 30;
    }

    if (!this.properties.lookAheadDays) {
      this.properties.lookAheadDays = 120;
    }

    if (!this.properties.calendarSources) {
      this.properties.calendarSources = '';
    }

    if (!this.properties.categoryColorMappings) {
      this.properties.categoryColorMappings = '';
    }

    if (!this.properties.swimLaneMappings) {
      this.properties.swimLaneMappings = '';
    }

    return Promise.resolve();
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.GeneralSettingsGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneSlider('lookBackDays', {
                  label: strings.LookBackDaysFieldLabel,
                  min: 0,
                  max: 180,
                  step: 5
                }),
                PropertyPaneSlider('lookAheadDays', {
                  label: strings.LookAheadDaysFieldLabel,
                  min: 30,
                  max: 365,
                  step: 5
                })
              ]
            },
            {
              groupName: strings.CalendarSourcesGroupName,
              groupFields: [
                PropertyPaneTextField('calendarSources', {
                  label: strings.CalendarSourcesFieldLabel,
                  description: strings.CalendarSourcesFieldDescription,
                  multiline: true,
                  rows: 8
                }),
                PropertyPaneTextField('swimLaneMappings', {
                  label: strings.SwimLaneMappingsFieldLabel,
                  description: strings.SwimLaneMappingsFieldDescription,
                  multiline: true,
                  rows: 6
                }),
                PropertyPaneTextField('categoryColorMappings', {
                  label: strings.CategoryColorMappingsFieldLabel,
                  description: strings.CategoryColorMappingsFieldDescription,
                  multiline: true,
                  rows: 6
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
