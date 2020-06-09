import { IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';
import ICalendarEvents from "../models/ICalendarEvents";

export default interface ISPOperations {
    LoadLists(): Promise<IDropdownOption[]>;
    LoadFields(fieldType: string, listTitle: string): Promise<IDropdownOption[]>;
    LoadEvents(listTitle: string, titleField: string, startDateField: string, endDateField: string, descField: string, allDayEventField: string, ShowRecurrenceEventsField: boolean): Promise<ICalendarEvents[]>;
}
