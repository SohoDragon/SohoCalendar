import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import * as strings from 'SpfxModernCalendarWebPartStrings';
import MCWCalendar from './components/MCWCalendar';
import { IMCWCalendarProps } from './models/IMCWCalendarProps';
import { PropertyPaneAsyncDropdown } from './components/PropertyPane/PropertyPaneAsyncDropdown';
import { IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';
import { update, get } from '@microsoft/sp-lodash-subset';
import ISpfxModernCalendarWebPartProps from './models/ISpfxModernCalendarWebPartProps';

import SPOperations from "./common/SPOperations";
import ISPOperations from "./models/ISPOperations";
import ICalendarEvents from "./models/ICalendarEvents";
import { defaultSelectKey } from './common/constants';

import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';

export default class SpfxModernCalendarWebPart extends BaseClientSideWebPart<ISpfxModernCalendarWebPartProps> {
  private ListDropDown: PropertyPaneAsyncDropdown;
  private StartDateDropDown: PropertyPaneAsyncDropdown;
  private EndDateDropDown: PropertyPaneAsyncDropdown;
  private EventTitleDropDown: PropertyPaneAsyncDropdown;
  private EventDescriptionDropDown: PropertyPaneAsyncDropdown;
  private AllDaysEventDropDown: PropertyPaneAsyncDropdown;

  private SPOperations: ISPOperations;

  private Events: ICalendarEvents[];

  protected async onInit(): Promise<void> {
    this.SPOperations = new SPOperations(
      this.context
    );

    this.Events = await this.SPOperations.loadEvents(this.properties.ListTitle, this.properties.EventTitleField,
      this.properties.StartDateField, this.properties.EndDateField, this.properties.EventDescriptionField,
      this.properties.AllDaysEventField);
  }

  public render(): void {
    const element: React.ReactElement<IMCWCalendarProps> = React.createElement(
      MCWCalendar,
      {
        context: this.context,
        WebpartTitle_compo: this.properties.WebPartTitle,
        EventBGColor_compo: this.properties.EventBGColor,
        EventTitleColor_compo: this.properties.EventTitleColor,
        ListTitle_compo: this.properties.ListTitle,
        StartDateField_compo: this.properties.StartDateField,
        EndDateField_compo: this.properties.EndDateField,
        EventTitleField_compo: this.properties.EventTitleField,
        EventDescriptionField_compo: this.properties.EventDescriptionField,
        AllDaysEventField_combo: this.properties.AllDaysEventField,
        Events: this.Events
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private loadFields(fieldType: string): Promise<IDropdownOption[]> {
    if (!this.properties.ListTitle) {
      // resolve to empty options since no list has been selected
      return Promise.resolve();
    }

    const wp: SpfxModernCalendarWebPart = this;

    return this.SPOperations.loadFields(fieldType, wp.properties.ListTitle);
  }

  private async onListFieldChange(propertyPath: string, newValue: any): Promise<void> {
    const oldValue: any = get(this.properties, propertyPath);
    // store new value in web part properties
    update(this.properties, propertyPath, (): any => { return newValue; });

    this.Events = [];

    this.Events = await this.SPOperations.loadEvents(this.properties.ListTitle, this.properties.EventTitleField,
      this.properties.StartDateField, this.properties.EndDateField, this.properties.EventDescriptionField,
      this.properties.AllDaysEventField);

    // refresh web part
    this.render();    
  }

  private async onListChange(propertyPath: string, newValue: any): Promise<void> {
    const oldValue: any = get(this.properties, propertyPath);
    // store new value in web part properties
    update(this.properties, propertyPath, (): any => { return newValue; });
    // reset selected item
    this.properties.StartDateField = defaultSelectKey;
    this.properties.EndDateField = defaultSelectKey;
    this.properties.EventTitleField = defaultSelectKey;
    this.properties.EventDescriptionField = defaultSelectKey;
    this.properties.AllDaysEventField = defaultSelectKey;

    // store new value in web part properties
    update(this.properties, 'StartDateField', (): any => { return this.properties.StartDateField; });
    update(this.properties, 'EndDateField', (): any => { return this.properties.EndDateField; });
    update(this.properties, 'EventTitleField', (): any => { return this.properties.EventTitleField; });
    update(this.properties, 'EventDescriptionField', (): any => { return this.properties.EventDescriptionField; });
    update(this.properties, 'AllDaysEventField', ():any => {return this.properties.AllDaysEventField; });
    this.Events = [];

    this.Events = await this.SPOperations.loadEvents(this.properties.ListTitle, this.properties.EventTitleField,
      this.properties.StartDateField, this.properties.EndDateField, this.properties.EventDescriptionField,
      this.properties.AllDaysEventField);

    // refresh web part
    this.render();
    // reset selected values in item dropdown
    this.StartDateDropDown.properties.selectedKey = this.properties.StartDateField;
    this.EndDateDropDown.properties.selectedKey = this.properties.EndDateField;
    this.EventTitleDropDown.properties.selectedKey = this.properties.EventTitleField;
    this.EventDescriptionDropDown.properties.selectedKey = this.properties.EventDescriptionField;
    this.AllDaysEventDropDown.properties.selectedKey = this.properties.AllDaysEventField;

    // allow to load items
    this.StartDateDropDown.properties.disabled = false;
    this.EndDateDropDown.properties.disabled = false;
    this.EventTitleDropDown.properties.disabled = false;
    this.EventDescriptionDropDown.properties.disabled = false;
    this.AllDaysEventDropDown.properties.disabled = false;
    
    // load items and re-render items dropdown
    this.StartDateDropDown.render();
    this.EndDateDropDown.render();
    this.EventTitleDropDown.render();
    this.EventDescriptionDropDown.render();
    this.AllDaysEventDropDown.render();

    this.context.propertyPane.refresh();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    this.ListDropDown = new PropertyPaneAsyncDropdown('ListTitle', {
      label: strings.ListTitleLabel,
      loadOptions: this.SPOperations.loadLists.bind(this),
      onPropertyChange: this.onListChange.bind(this),
      selectedKey: this.properties.ListTitle
    });

    // reference to item dropdown needed later after selecting a list
    this.StartDateDropDown = new PropertyPaneAsyncDropdown('StartDateField', {
      label: strings.StartDateFieldLabel,
      loadOptions: this.loadFields.bind(this, "DateTime"),
      onPropertyChange: this.onListFieldChange.bind(this),
      selectedKey: this.properties.StartDateField,
      // should be disabled if no list has been selected
      disabled: !this.properties.ListTitle
    });

    // reference to item dropdown needed later after selecting a list
    this.EndDateDropDown = new PropertyPaneAsyncDropdown('EndDateField', {
      label: strings.EndDateFieldLabel,
      loadOptions: this.loadFields.bind(this, "DateTime"),
      onPropertyChange: this.onListFieldChange.bind(this),
      selectedKey: this.properties.EndDateField,
      // should be disabled if no list has been selected
      disabled: !this.properties.ListTitle
    });

    // reference to item dropdown needed later after selecting a list
    this.EventTitleDropDown = new PropertyPaneAsyncDropdown('EventTitleField', {
      label: strings.EventTitleFieldLabel,
      loadOptions: this.loadFields.bind(this, "Text"),
      onPropertyChange: this.onListFieldChange.bind(this),
      selectedKey: this.properties.EventTitleField,
      // should be disabled if no list has been selected
      disabled: !this.properties.ListTitle
    });

    // reference to item dropdown needed later after selecting a list
    this.EventDescriptionDropDown = new PropertyPaneAsyncDropdown('EventDescriptionField', {
      label: strings.EventDescriptionFieldLabel,
      loadOptions: this.loadFields.bind(this, "Note"),
      onPropertyChange: this.onListFieldChange.bind(this),
      selectedKey: this.properties.EventDescriptionField,
      // should be disabled if no list has been selected
      disabled: !this.properties.ListTitle
    });

    this.AllDaysEventDropDown = new PropertyPaneAsyncDropdown('AllDaysEventField', {
      label: strings.AllDaysEventFieldLabel,
      loadOptions: this.loadFields.bind(this, "AllDayEvent"),
      onPropertyChange: this.onListFieldChange.bind(this),
      selectedKey: this.properties.AllDaysEventField,
      // should be disabled if no list has been selected
      disabled: !this.properties.ListTitle
    });

    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('WebPartTitle', {
                  label: strings.WebPartTitleLabel
                }),
                PropertyFieldColorPicker('EventBGColor', {
                  label: strings.EventBackgroundColor,
                  selectedColor: this.properties.EventBGColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Inline,
                  iconName: 'Precipitation',
                  key: 'EventBGColor'
                }),
                PropertyFieldColorPicker('EventTitleColor', {
                  label: strings.EventTitleColor,
                  selectedColor: this.properties.EventTitleColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Inline,
                  iconName: 'Precipitation',
                  key: 'EventTitleColor'
                })
              ]
            },
            {
              groupName: strings.ListInfoGroupName,
              groupFields: [
                this.ListDropDown,
                this.StartDateDropDown,
                this.EndDateDropDown,
                this.EventTitleDropDown,
                this.EventDescriptionDropDown,
                this.AllDaysEventDropDown
              ]
            }
          ]
        }
      ]
    };
  }
}
