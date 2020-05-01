import { IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';
import ICalendarEvents from "../models/ICalendarEvents";

export default interface ISPOperations {
    loadLists(): Promise<IDropdownOption[]>;
    loadFields(fieldType: string, listTitle: string): Promise<IDropdownOption[]>;
    loadEvents(listTitle: string, titleField: string, startDateField: string, endDateField: string, descField: string, allDayEventField: string): Promise<ICalendarEvents[]>;
}
