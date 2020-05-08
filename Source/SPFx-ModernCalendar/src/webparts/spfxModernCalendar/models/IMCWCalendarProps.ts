import { WebPartContext } from "@microsoft/sp-webpart-base";
import ICalendarEvents from "../models/ICalendarEvents";

export interface IMCWCalendarProps {
  context: WebPartContext;
  WebpartTitle_compo: string;
  EventBGColor_compo: string;
  EventTitleColor_compo: string;
  ListTitle_compo: string;
  StartDateField_compo: string;
  EndDateField_compo: string;
  EventTitleField_compo: string;
  EventDescriptionField_compo: string;
  AllDaysEventField_combo: string;
  ShowRecurrenceEventsField_combo: boolean;
  Events: ICalendarEvents[];
}