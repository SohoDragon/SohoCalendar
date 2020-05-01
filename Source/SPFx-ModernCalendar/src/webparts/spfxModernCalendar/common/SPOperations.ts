import ISPOperations from "../models/ISPOperations";
import ICalendarEvents from "../models/ICalendarEvents";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';
import { defaultSelectKey, defaultSelectText } from '../common/constants';

import { sp } from "@pnp/sp/presets/all";

export default class SPOperations implements ISPOperations {
    private context: WebPartContext;

    constructor(context: WebPartContext) {
        this.context = context;

        sp.setup({
            spfxContext: this.context
        });
    }

    public loadLists(): Promise<IDropdownOption[]> {
        return new Promise<IDropdownOption[]>((resolve: (options: IDropdownOption[]) => void, reject: (error: any) => void) => {
            sp.web.lists.filter('Hidden eq false').get().then((data) => {
                let listArr = [];

                listArr.push({
                    key: defaultSelectKey,
                    text: defaultSelectText
                });

                data.map((item, key) => {
                    listArr.push({
                        key: item.Title,
                        text: item.Title
                    });
                });
                resolve(listArr);
            }).catch((err) => {
                console.log("Load Lists err: " + err);

                let listArr = [];
                listArr.push({
                    key: defaultSelectKey,
                    text: defaultSelectText
                });
                resolve(listArr);
            });
        });
    }

    public loadFields(fieldType: string, listTitle: string): Promise<IDropdownOption[]> {
        return new Promise<IDropdownOption[]>((resolve: (options: IDropdownOption[]) => void, reject: (error: any) => void) => {
            let filter = "";

            if (fieldType !== "AllDayEvent") {
                filter = "Hidden eq false and ReadOnlyField eq false and TypeAsString eq '" + fieldType + "'";
            }
            else {
                filter = "Hidden eq false and ReadOnlyField eq false and (TypeAsString eq '" + fieldType + "' or TypeAsString eq 'Boolean')";
            }

            sp.web.lists.getByTitle(listTitle).fields.select('Title, InternalName, TypeAsString').filter(filter).get().then((data) => {
                let listFieldArr = [];

                listFieldArr.push({
                    key: defaultSelectKey,
                    text: defaultSelectText
                });

                data.map((item, key) => {
                    listFieldArr.push({
                        key: item.InternalName,
                        text: item.Title
                    });
                });
                resolve(listFieldArr);
            }).catch((err) => {
                console.log("Load Field err: " + err);
                reject(err);
            });
        });
    }

    public loadEvents(listTitle: string, titleField: string, startDateField: string, endDateField: string, descField: string, allDayEventField: string): Promise<ICalendarEvents[]> {
        return new Promise<ICalendarEvents[]>((resolve: (events: ICalendarEvents[]) => void, reject: (error: any) => void) => {
            sp.web.lists.getByTitle(listTitle).items.get().then((data) => {
                let itemArr = [];
                if (listTitle && titleField && startDateField && endDateField && descField &&
                    listTitle != defaultSelectKey && titleField != defaultSelectKey && 
                    startDateField != defaultSelectKey && endDateField != defaultSelectKey && 
                    descField != defaultSelectKey) {
                    data.map((item, key) => {
                        //We are ignore the recurrence events as this is not getting the proper results.
                        if (!item.fRecurrence) {
                            itemArr.push({
                                id: item.Id,
                                title: item[titleField],
                                start: new Date(item[startDateField]),
                                end: new Date(item[endDateField]),
                                desc: item[descField] ? item[descField] : "",
                                allDay: item[allDayEventField] ? item[allDayEventField] : false
                            });
                        }
                    });
                }
                resolve(itemArr);
            }).catch((err) => {
                console.log("Load Events err: " + err);
                reject(err);
            });
        });
    }
}