import ISPOperations from "../models/ISPOperations";
import ICalendarEvents from "../models/ICalendarEvents";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';
import { defaultSelectKey, defaultSelectText, weekDays, sunday, monday, tuesday, wednesday, thrusday, friday, saturday } from '../common/constants';

import pnp from "sp-pnp-js";
import * as moment from "moment";

export default class SPOperations implements ISPOperations {
    private context: WebPartContext;

    constructor(context: WebPartContext) {
        this.context = context;

        pnp.setup({
            spfxContext: this.context
        });
    }

    public LoadLists(): Promise<IDropdownOption[]> {
        return new Promise<IDropdownOption[]>((resolve: (options: IDropdownOption[]) => void, reject: (error: any) => void) => {
            pnp.sp.web.lists.filter('Hidden eq false').get().then((data) => {
                let listArr = [];

                listArr.push({
                    key: defaultSelectKey,
                    text: defaultSelectText
                });

                data.map((item, key) => {
                    listArr.push({
                        key: item.Title + ";#" + item.BaseTemplate,
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

    public LoadFields(fieldType: string, listTitle: string): Promise<IDropdownOption[]> {
        listTitle = listTitle.split(";#")[0];
        return new Promise<IDropdownOption[]>((resolve: (options: IDropdownOption[]) => void, reject: (error: any) => void) => {
            let filter = "";

            if (fieldType !== "AllDayEvent") {
                filter = "Hidden eq false and ReadOnlyField eq false and TypeAsString eq '" + fieldType + "'";
            }
            else {
                filter = "Hidden eq false and ReadOnlyField eq false and (TypeAsString eq '" + fieldType + "' or TypeAsString eq 'Boolean')";
            }

            pnp.sp.web.lists.getByTitle(listTitle).fields.select('Title, InternalName, TypeAsString').filter(filter).get().then((data) => {
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

    public async LoadEvents(listTitle: string, titleField: string, startDateField: string, endDateField: string, descField: string, allDayEventField: string, ShowRecurrenceEventsField: boolean): Promise<ICalendarEvents[]> {
        listTitle = listTitle.split(";#")[0];
        return new Promise<ICalendarEvents[]>(async (resolve: (events: ICalendarEvents[]) => void, reject: (error: any) => void) => {
            let itemArr = [];
            let removeItemArr = [];
            let editItemArr = [];
            if (listTitle && titleField && startDateField && endDateField && descField &&
                listTitle != defaultSelectKey && titleField != defaultSelectKey &&
                startDateField != defaultSelectKey && endDateField != defaultSelectKey &&
                descField != defaultSelectKey) {

                let selectFields = "";

                //Recurring Events only works in the Calendar type list. 
                if (ShowRecurrenceEventsField) {
                    selectFields = "Id, " + titleField + ", " + startDateField + ", " + endDateField + ", " + descField + ", fRecurrence, RecurrenceData, MasterSeriesItemID, RecurrenceID, FieldValuesAsText/"+startDateField+", FieldValuesAsText/"+endDateField+"";
                }

                let dataArray = [];
                await pnp.sp.web.lists.getByTitle(listTitle).items.select(selectFields).expand("FieldValuesAsText").getPaged().then(async (data) => {
                    dataArray = dataArray.concat(data.results);
                    while (data["nextUrl"]) {
                        await data.getNext().then((d) => {
                            dataArray = dataArray.concat(d.results);
                            data = d;
                        }).catch((err) => {
                            reject(err);
                        });
                    }
                }).catch((err) => {
                    reject(err);
                });

                console.log(dataArray);
                
                // pnp.sp.web.lists.getByTitle(listTitle).items.select(selectFields).get().then((data) => {
                Promise.all(dataArray.map(async (item, key) => {
                    if (!item.fRecurrence) {
                        let regStartDate;
                        let regEndDate;

                        if (!item[allDayEventField]) {
                            //regStartDate = await pnp.sp.web.regionalSettings.timeZone.utcToLocalTime((new Date(item[startDateField])));
                            //regEndDate = await pnp.sp.web.regionalSettings.timeZone.utcToLocalTime((new Date(item[endDateField])));
                            //console.log(moment(item.FieldValuesAsText[startDateField]).format("YYYY-MM-DDTHH:mm:ss"), regStartDate);
                            regStartDate = moment(new Date(item.FieldValuesAsText[startDateField])).format("YYYY-MM-DDTHH:mm:ss");
                            regEndDate = moment(new Date(item.FieldValuesAsText[endDateField])).format("YYYY-MM-DDTHH:mm:ss");
                        }
                        else {
                            //let tempStartDate = new Date(item[startDateField]);
                            //let tempEndDate = new Date(item[endDateField]);
                            //regStartDate = await pnp.sp.web.regionalSettings.timeZone.utcToLocalTime((new Date(tempStartDate.setDate(tempStartDate.getDate() + 1))));
                            //regEndDate = await pnp.sp.web.regionalSettings.timeZone.utcToLocalTime((new Date(tempEndDate.setDate(tempEndDate.getDate() + 1))));
                            //console.log(regStartDate, moment(item.FieldValuesAsText[startDateField]).add('days', 1).format("YYYY-MM-DDTHH:mm:ss"));
                             regStartDate = moment(new Date(item.FieldValuesAsText[startDateField])).add(1, 'days').format("YYYY-MM-DDTHH:mm:ss");
                             regEndDate = moment(new Date(item.FieldValuesAsText[endDateField])).add(1, 'days').format("YYYY-MM-DDTHH:mm:ss");
                        }
                        itemArr.push({
                            id: item.Id,
                            recurrenceId: "",
                            title: item[titleField],
                            start: regStartDate,
                            end: regEndDate,
                            desc: item[descField] ? item[descField] : "",
                            allDay: item[allDayEventField] ? item[allDayEventField] : false
                        });
                    }
                    else if (ShowRecurrenceEventsField) {

                        let recurreceData = item["RecurrenceData"];

                        if (recurreceData && recurreceData.indexOf("<recurrence>") > -1) {
                            let recurringDataArr = await this.ProcessRecurringEvents(item, titleField, startDateField, endDateField, descField, allDayEventField);
                            itemArr = itemArr.concat(recurringDataArr);
                        }
                        else if (recurreceData && (item[titleField]).indexOf("Deleted:") > -1) {
                            //Add the remove events logic
                            removeItemArr.push({
                                RecurrenceID: item["RecurrenceID"],
                                MasterSeriesItemID: item["MasterSeriesItemID"]
                            });
                        }
                        else if (recurreceData) {
                            editItemArr.push(item);
                        }
                    }
                })).then((_) => {
                    if (editItemArr.length > 0 && itemArr.length > 0) {
                        itemArr = this.EditRecurringEvents(itemArr, editItemArr, titleField, descField, startDateField, endDateField);
                    }
                    if (removeItemArr.length > 0 && itemArr.length > 0) {
                        itemArr = this.RemoveDeletedEvents(itemArr, removeItemArr);
                    }
                    resolve(itemArr);
                }).catch((err) => {
                    console.log("Load Events Loop err: " + err);
                    if (editItemArr.length > 0 && itemArr.length > 0) {
                        itemArr = this.EditRecurringEvents(itemArr, editItemArr, titleField, descField, startDateField, endDateField);
                    }
                    if (removeItemArr.length > 0 && itemArr.length > 0) {
                        itemArr = this.RemoveDeletedEvents(itemArr, removeItemArr);
                    }
                    resolve(itemArr);
                });
                // }).catch((err) => {
                //     console.log("Load Events err: " + err);
                //     if (editItemArr.length > 0 && itemArr.length > 0) {
                //         itemArr = this.EditRecurringEvents(itemArr, editItemArr, titleField, descField, startDateField, endDateField);
                //     }
                //     if (removeItemArr.length > 0 && itemArr.length > 0) {
                //         itemArr = this.RemoveDeletedEvents(itemArr, removeItemArr);
                //     }
                //     resolve(itemArr);
                // });
            }
            else {
                resolve(itemArr);
            }
        });
    }

    private SetTime(date: Date, setTimeDate: Date): Date {

        let newTempDate = date;

        newTempDate.setHours(setTimeDate.getHours());
        newTempDate.setMinutes(setTimeDate.getMinutes());
        newTempDate.setSeconds(setTimeDate.getSeconds());
        newTempDate.setMilliseconds(setTimeDate.getMilliseconds());

        return newTempDate;
    }

    /* 
    For Last DayName, isLastDayName = true and LastDayNumber = number (e.g., Sunday=0, Monday=1, Tuesday=2, Wednesday=3, Thrusday=4, Friday=5, Saturday=6)
    For Last Day, isLastDay = true
    For Last WeekDay, isLastWeekDay = true
    For Last WeekEnd, isLastWeekEndDay = true
    */
    private GetLastDay(date: Date, dayWeekDayIndexArr: any, isLastDay: boolean, isLastDayName: boolean, lastDayNumber: Number, isLastWeekDay: boolean, isLastWeekEndDay: boolean): Date {
        let newDate = new Date(date.getTime());

        // Roll to the first day of month
        newDate.setDate(1);

        //Set the Next Month
        newDate.setMonth(newDate.getMonth() + 1);

        // Roll the days backwards until LastDayNumber
        if (isLastDay) {
            newDate.setDate(newDate.getDate() - 1);
        }
        else if (isLastWeekDay || isLastWeekEndDay) {
            let dayNumber;
            do {
                newDate.setDate(newDate.getDate() - 1);
                let dayName = weekDays[newDate.getDay()];
                let temp = dayWeekDayIndexArr.filter((el) => { return el[dayName] != null; });
                dayNumber = temp[0][dayName.toString()];
            } while (((dayNumber === 0 || dayNumber === 6) && isLastWeekDay) || (dayNumber > 0 && dayNumber < 6 && isLastWeekEndDay));

        }
        else {
            do {
                newDate.setDate(newDate.getDate() - 1);
            } while ((isLastDayName && newDate.getDay() !== lastDayNumber));
        }
        return newDate;
    }

    /*
    ApperanceNumber = number (e.g., first=1, second=2, etc...)
    For Appereance Day by day name, isDayName = true and dayNumber = number (e.g., Sunday=0, Monday=1, Tuesday=2, Wednesday=3, Thrusday=4, Friday=5, Saturday=6)
    For Appereance Day by day number, isDay = true
    For Last WeekDay, isWeekDay = true
    For Last WeekEnd, isWeekEndDay = true 
    */
    private GetRequiredDay(date: Date, apperanceNumber: any, dayWeekDayIndexArr: any, isDay: boolean, isDayName: boolean, dayNumber: number, isWeekDay: boolean, isWeekEndDay: boolean): Date {
        var newDate = new Date(date.getTime());

        // Roll to the first day of month
        newDate.setDate(1);

        if (isDay) {
            newDate.setDate(apperanceNumber);
        }
        else if (isDayName) {
            if (newDate.getDay() > 3) {
                newDate.setDate(newDate.getDate() + ((7 * apperanceNumber) + dayNumber) - newDate.getDay());
            }
            else {
                newDate.setDate(newDate.getDate() + ((7 * apperanceNumber) + dayNumber) - newDate.getDay() - 7);
            }
        }
        else if (isWeekDay) {
            var weekDayCount = 0;
            do {
                var dayNameWD = weekDays[newDate.getDay()];
                var tempWD = dayWeekDayIndexArr.filter((el) => { return el[dayNameWD] != null; });
                dayNumber = tempWD[0][dayNameWD.toString()];
                if (dayNumber > 0 && dayNumber < 6) {
                    weekDayCount++;
                }
                if (weekDayCount !== apperanceNumber) {
                    newDate.setDate(newDate.getDate() + 1);
                }
            } while (weekDayCount !== apperanceNumber);
        }
        else if (isWeekEndDay) {
            var weekEndCount = 0;
            do {
                var dayNameWED = weekDays[newDate.getDay()];
                var tempWED = dayWeekDayIndexArr.filter((el) => { return el[dayNameWED] != null; });
                dayNumber = tempWED[0][dayNameWED.toString()];
                if (dayNumber === 0 || dayNumber === 6) {
                    weekEndCount++;
                }
                if (weekEndCount !== apperanceNumber) {
                    newDate.setDate(newDate.getDate() + 1);
                }
            } while (weekEndCount !== apperanceNumber);
        }
        else {
            return null;
        }

        if (date.getMonth() == newDate.getMonth()) {
            return newDate;
        }
        else {
            return null;
        }
    }

    private ProcessRecurringEvents(item: any, titleField: string, startDateField: string, endDateField: string, descField: string, allDayEventField: string): Promise<ICalendarEvents[]> {
        return new Promise<ICalendarEvents[]>((resolve: (events: ICalendarEvents[]) => void, reject: (error: any) => void) => {
            let itemArr = [];
            let recurreceData = item["RecurrenceData"];

            try {
                if (recurreceData) {
                    let eventStartDate = new Date(item[startDateField]);
                    let eventEndDate = new Date(item[endDateField]);

                    //If the firstDayOfWeek = su then this is the default Array.
                    let DayWeekDayIndexArr = [{ su: 0 }, { mo: 1 }, { tu: 2 }, { we: 3 }, { th: 4 }, { fr: 5 }, { sa: 6 }];

                    const parser = new DOMParser();
                    const xml = parser.parseFromString(recurreceData, 'text/xml');

                    let MonthlyRecurrenceArr = [];
                    let YearlyRecurrenceArr = [];

                    let daily = xml.querySelector('daily');
                    let weekly = xml.querySelector('weekly');
                    let monthly = xml.querySelector('monthly') || xml.querySelector('monthlyByDay');
                    let yearly = xml.querySelector('yearly') || xml.querySelector('yearlyByDay');

                    let firstDayOfWeek = xml.querySelector('firstDayOfWeek').textContent;

                    if (firstDayOfWeek !== sunday) {
                        let index = weekDays.indexOf(firstDayOfWeek);
                        let startIndex = 0;

                        for (let k = index; k < DayWeekDayIndexArr.length; k++) {
                            DayWeekDayIndexArr[k][weekDays[k]] = startIndex;
                            startIndex++;
                        }
                        for (let l = 0; l < index; l++) {
                            DayWeekDayIndexArr[l][weekDays[l]] = startIndex;
                            startIndex++;
                        }
                    }

                    //#region Daily Events Process
                    if (daily) {
                        let dayFrequency = daily.getAttribute('dayFrequency');
                        let weekday = daily.getAttribute('weekday');

                        if (dayFrequency) {
                            //#region XML Sample
                            /* <recurrence>
                            <rule>
                                <firstDayOfWeek>su</firstDayOfWeek>
                                <repeat>
                                    <daily dayFrequency="1" />
                                </repeat>
                                <repeatForever>FALSE</repeatForever>
                            </rule>
                            </recurrence>
                            <recurrence>
                            <rule>
                                <firstDayOfWeek>su</firstDayOfWeek>
                                <repeat>
                                    <daily dayFrequency="1" />
                                </repeat>
                                <windowEnd>2020-05-31T10:00:00Z</windowEnd>
                            </rule>
                            </recurrence> */
                            //#endregion

                            for (let i = eventStartDate; i <= eventEndDate; i.setDate(i.getDate() + parseInt(dayFrequency))) {

                                let newTempStartDate = new Date(i.getTime());
                                let newTempEndDate = new Date(i.getTime());

                                let Startdate = this.SetTime(newTempStartDate, eventStartDate);
                                let EndDate = this.SetTime(newTempEndDate, eventEndDate);

                                itemArr.push({
                                    id: item.Id,
                                    recurrenceId: item.Id + ".0." + Startdate.toISOString().split('.')[0] + "Z",
                                    title: item[titleField],
                                    start: Startdate,
                                    end: EndDate,
                                    desc: item[descField] ? item[descField] : "",
                                    allDay: item[allDayEventField] ? item[allDayEventField] : false
                                });
                            }
                        }
                        else if (weekday) {
                            //#region XML Sample
                            /* <recurrence>
                               <rule>
                                  <firstDayOfWeek>su</firstDayOfWeek>
                                  <repeat>
                                     <daily weekday="TRUE" />
                                  </repeat>
                                  <repeatInstances>10</repeatInstances>
                               </rule>
                            </recurrence>
                            
                            <recurrence>
                               <rule>
                                  <firstDayOfWeek>su</firstDayOfWeek>
                                  <repeat>
                                     <daily weekday="TRUE" />
                                  </repeat>
                                  <repeatInstances>10</repeatInstances>
                               </rule>
                            </recurrence>
                            <recurrence>
                               <rule>
                                  <firstDayOfWeek>su</firstDayOfWeek>
                                  <repeat>
                                     <daily weekday="TRUE" />
                                  </repeat>
                                  <windowEnd>2020-05-31T10:00:00Z</windowEnd>
                               </rule>
                            </recurrence> */
                            //#endregion

                            for (let j = eventStartDate; j <= eventEndDate; j.setDate(j.getDate() + 1)) {

                                let newTempWeekDayStartDate = new Date(j.getTime());
                                let newTempWeekDayEndDate = new Date(j.getTime());

                                let dayName = weekDays[newTempWeekDayStartDate.getDay()];

                                const temp = DayWeekDayIndexArr.filter((el) => { return el[dayName] != null; });

                                let dayNumber = temp[0][dayName.toString()];

                                //Check is WeekDay or not
                                if (dayNumber > 0 && dayNumber < 6) {
                                    let Startdate1 = this.SetTime(newTempWeekDayStartDate, eventStartDate);
                                    let EndDate1 = this.SetTime(newTempWeekDayEndDate, eventEndDate);

                                    itemArr.push({
                                        id: item.Id,
                                        recurrenceId: item.Id + ".0." + Startdate1.toISOString().split('.')[0] + "Z",
                                        title: item[titleField],
                                        start: Startdate1,
                                        end: EndDate1,
                                        desc: item[descField] ? item[descField] : "",
                                        allDay: item[allDayEventField] ? item[allDayEventField] : false
                                    });
                                }
                            }
                        }
                    }
                    //#endregion
                    //#region Weekly Events Process
                    else if (weekly) {

                        //#region XML Sample
                        /*<recurrence>
                           <rule>
                              <firstDayOfWeek>su</firstDayOfWeek>
                              <repeat>
                                 <weekly tu="TRUE" th="TRUE" weekFrequency="1" />
                              </repeat>
                              <repeatForever>FALSE</repeatForever>
                           </rule>
                        </recurrence>
                        <recurrence>
                           <rule>
                              <firstDayOfWeek>su</firstDayOfWeek>
                              <repeat>
                                 <weekly mo="TRUE" tu="TRUE" weekFrequency="1" />
                              </repeat>
                              <repeatForever>FALSE</repeatForever>
                           </rule>
                        </recurrence>
                        <recurrence>
                           <rule>
                              <firstDayOfWeek>su</firstDayOfWeek>
                              <repeat>
                                 <weekly su="TRUE" mo="TRUE" tu="TRUE" fr="TRUE" sa="TRUE" weekFrequency="2" />
                              </repeat>
                              <repeatInstances>10</repeatInstances>
                           </rule>
                        </recurrence>
                        <recurrence>
                           <rule>
                              <firstDayOfWeek>su</firstDayOfWeek>
                              <repeat>
                                 <weekly su="TRUE" mo="TRUE" tu="TRUE" sa="TRUE" weekFrequency="1" />
                              </repeat>
                              <windowEnd>2020-05-29T10:00:00Z</windowEnd>
                           </rule>
                        </recurrence>*/
                        //#endregion

                        let weekFrequency = weekly.getAttribute('weekFrequency');
                        let isSunday = weekly.getAttribute(sunday);
                        let isMonday = weekly.getAttribute(monday);
                        let isTuesday = weekly.getAttribute(tuesday);
                        let isWednesday = weekly.getAttribute(wednesday);
                        let isThrusday = weekly.getAttribute(thrusday);
                        let isFriday = weekly.getAttribute(friday);
                        let isSaturday = weekly.getAttribute(saturday);

                        let WeekCount = 0;
                        let isSkipped = false;
                        let eventNumber = 0;

                        for (let i = eventStartDate; i <= eventEndDate; i.setDate(i.getDate() + 1)) {

                            let newTempStartDate = new Date(i.getTime());
                            let newTempEndDate = new Date(i.getTime());

                            let Startdate = this.SetTime(newTempStartDate, eventStartDate);
                            let EndDate = this.SetTime(newTempEndDate, eventEndDate);

                            let dayName = weekDays[newTempStartDate.getDay()];
                            // let isNewWeekStarted = false;
                            if (dayName == firstDayOfWeek && eventNumber != 0 &&
                                isSkipped === false && parseInt(weekFrequency) > 1) {
                                let skipdays = ((parseInt(weekFrequency) - 1) * 7) - 1;
                                i.setDate(i.getDate() + skipdays);
                                isSkipped = true;
                                continue;
                            }
                            eventNumber++;
                            isSkipped = false;

                            if ((dayName === sunday && isSunday) || (dayName === monday && isMonday) || (dayName === tuesday && isTuesday)
                                || (dayName === wednesday && isWednesday) || (dayName === thrusday && isThrusday) || (dayName === friday && isFriday)
                                || (dayName === saturday && isSaturday)) {
                                itemArr.push({
                                    id: item.Id,
                                    recurrenceId: item.Id + ".0." + Startdate.toISOString().split('.')[0] + "Z",
                                    title: item[titleField],
                                    start: Startdate,
                                    end: EndDate,
                                    desc: item[descField] ? item[descField] : "",
                                    allDay: item[allDayEventField] ? item[allDayEventField] : false
                                });
                            }
                        }
                    }
                    //#endregion
                    //#region Monthly Events Process
                    else if (monthly) {
                        //Currently this is only for info
                        MonthlyRecurrenceArr.push({
                            data: recurreceData
                        });

                        let attrMonthly = xml.querySelector('monthly');
                        let attrMonthlyByDay = xml.querySelector('monthlyByDay');
                        let monthFrequency = monthly.getAttribute('monthFrequency');

                        let tempEventStartDate = new Date(eventStartDate.getTime());
                        let tempEventEndDate = new Date(eventEndDate.getTime());

                        if (attrMonthly) {

                            //#region XML sample
                            /* <recurrence>
                                <rule>
                                    <firstDayOfWeek>su</firstDayOfWeek>
                                    <repeat>
                                        <monthly monthFrequency="1" day="5" />
                                    </repeat>
                                    <repeatForever>FALSE</repeatForever>
                                </rule>
                            </recurrence> */
                            //#endregion

                            let day = attrMonthly.getAttribute('day');

                            for (let i = eventStartDate; i <= eventEndDate; i = new Date(new Date(i.setMonth(i.getMonth() + parseInt(monthFrequency))).setDate(parseInt(day)))) {

                                let newTempStartDate = new Date(i.getTime());
                                let newTempEndDate = new Date(i.getTime());

                                newTempStartDate.setDate(parseInt(day));
                                newTempEndDate.setDate(parseInt(day));

                                let Startdate = this.SetTime(newTempStartDate, eventStartDate);
                                let EndDate = this.SetTime(newTempEndDate, eventEndDate);

                                if (Startdate >= tempEventStartDate && Startdate <= tempEventEndDate) {
                                    itemArr.push({
                                        id: item.Id,
                                        recurrenceId: item.Id + ".0." + Startdate.toISOString().split('.')[0] + "Z",
                                        title: item[titleField],
                                        start: Startdate,
                                        end: EndDate,
                                        desc: item[descField] ? item[descField] : "",
                                        allDay: item[allDayEventField] ? item[allDayEventField] : false
                                    });
                                }
                            }
                        }
                        else if (attrMonthlyByDay) {

                            //#region XML Sample
                            /* <recurrence>
                                <rule>
                                    <firstDayOfWeek>su</firstDayOfWeek>
                                    <repeat>
                                        <monthlyByDay we="TRUE" weekdayOfMonth="second" monthFrequency="2" />
                                    </repeat>
                                    <repeatForever>FALSE</repeatForever>
                                </rule>
                            </recurrence>
                            <recurrence>
                                <rule>
                                    <firstDayOfWeek>su</firstDayOfWeek>
                                    <repeat>
                                        <monthlyByDay su="TRUE" weekdayOfMonth="first" monthFrequency="2" />
                                    </repeat>
                                    <repeatInstances>10</repeatInstances>
                                </rule>
                            </recurrence>
                            <recurrence>
                                <rule>
                                    <firstDayOfWeek>su</firstDayOfWeek>
                                    <repeat>
                                        <monthlyByDay day="TRUE" weekdayOfMonth="first" monthFrequency="1" />
                                    </repeat>
                                    <windowEnd>2020-06-03T15:00:00Z</windowEnd>
                                </rule>
                            </recurrence>
                            <recurrence>
                                <rule>
                                    <firstDayOfWeek>su</firstDayOfWeek>
                                    <repeat>
                                        <monthlyByDay weekday="TRUE" weekdayOfMonth="first" monthFrequency="1" />
                                    </repeat>
                                    <windowEnd>2020-06-03T15:00:00Z</windowEnd>
                                </rule>
                            </recurrence>
                            <recurrence>
                                <rule>
                                    <firstDayOfWeek>su</firstDayOfWeek>
                                    <repeat>
                                        <monthlyByDay weekend_day="TRUE" weekdayOfMonth="first" monthFrequency="1" />
                                    </repeat>
                                    <windowEnd>2020-06-08T15:00:00Z</windowEnd>
                                </rule>
                            </recurrence>
                            <recurrence>
                                <rule>
                                    <firstDayOfWeek>su</firstDayOfWeek>
                                    <repeat>
                                        <monthlyByDay sa="TRUE" weekdayOfMonth="first" monthFrequency="1" />
                                    </repeat>
                                    <windowEnd>2020-07-30T10:00:00Z</windowEnd>
                                </rule>
                            </recurrence>
                            <recurrence>
                                <rule>
                                    <firstDayOfWeek>su</firstDayOfWeek>
                                    <repeat>
                                        <monthlyByDay weekday="TRUE" weekdayOfMonth="first" monthFrequency="1" />
                                    </repeat>
                                    <repeatForever>FALSE</repeatForever>
                                </rule>
                            </recurrence> */
                            //#endregion

                            let weekdayOfMonth = attrMonthlyByDay.getAttribute('weekdayOfMonth');
                            let isSunday = attrMonthlyByDay.getAttribute(sunday);
                            let isMonday = attrMonthlyByDay.getAttribute(monday);
                            let isTuesday = attrMonthlyByDay.getAttribute(tuesday);
                            let isWednesday = attrMonthlyByDay.getAttribute(wednesday);
                            let isThrusday = attrMonthlyByDay.getAttribute(thrusday);
                            let isFriday = attrMonthlyByDay.getAttribute(friday);
                            let isSaturday = attrMonthlyByDay.getAttribute(saturday);
                            let day = attrMonthlyByDay.getAttribute("day");
                            let weekday = attrMonthlyByDay.getAttribute("weekday");
                            let weekend_day = attrMonthlyByDay.getAttribute("weekend_day");

                            let weekDayNumber = isSunday ? 0 : (
                                isMonday ? 1 : (
                                    isTuesday ? 2 : (
                                        isWednesday ? 3 : (
                                            isThrusday ? 4 : (
                                                isFriday ? 5 : (
                                                    isSaturday ? 6 : 0
                                                )
                                            )
                                        )
                                    )
                                )
                            );

                            for (let i = eventStartDate; i <= eventEndDate; i.setMonth(i.getMonth() + parseInt(monthFrequency))) {
                                let newTempStartDate = new Date(i.getTime());
                                let newTempEndDate = new Date(i.getTime());

                                let fnEventStartDate;
                                let fnEventEndDate;

                                if (weekdayOfMonth === "last") {
                                    if (day) {
                                        fnEventStartDate = this.GetLastDay(newTempStartDate, DayWeekDayIndexArr, true, false, 0, false, false);
                                    }
                                    else if (weekday) {
                                        fnEventStartDate = this.GetLastDay(newTempStartDate, DayWeekDayIndexArr, false, false, 0, true, false);
                                    }
                                    else if (weekend_day) {
                                        fnEventStartDate = this.GetLastDay(newTempStartDate, DayWeekDayIndexArr, false, false, 0, false, true);
                                    }
                                    else if (isSunday || isMonday || isTuesday || isWednesday || isThrusday || isFriday || isSaturday) {
                                        fnEventStartDate = this.GetLastDay(newTempStartDate, DayWeekDayIndexArr, false, true, weekDayNumber, false, false);
                                    }
                                }
                                else if (weekdayOfMonth) {
                                    let apperanceNumber = weekdayOfMonth === "first" ? 1 : (
                                        weekdayOfMonth === "second" ? 2 : (
                                            weekdayOfMonth === "third" ? 3 : (
                                                weekdayOfMonth === "fourth" ? 4 : 0
                                            )
                                        )
                                    );

                                    if (day) {
                                        fnEventStartDate = this.GetRequiredDay(newTempStartDate, apperanceNumber, DayWeekDayIndexArr, true, false, 0, false, false);
                                    }
                                    else if (weekday) {
                                        fnEventStartDate = this.GetRequiredDay(newTempStartDate, apperanceNumber, DayWeekDayIndexArr, false, false, 0, true, false);
                                    }
                                    else if (weekend_day) {
                                        fnEventStartDate = this.GetRequiredDay(newTempStartDate, apperanceNumber, DayWeekDayIndexArr, false, false, 0, false, true);
                                    }
                                    else if (isSunday || isMonday || isTuesday || isWednesday || isThrusday || isFriday || isSaturday) {
                                        fnEventStartDate = this.GetRequiredDay(newTempStartDate, apperanceNumber, DayWeekDayIndexArr, false, true, weekDayNumber, false, false);
                                    }
                                }

                                if (fnEventStartDate && fnEventStartDate !== null) {
                                    fnEventEndDate = new Date(fnEventStartDate.getTime());

                                    let Startdate = this.SetTime(fnEventStartDate, eventStartDate);
                                    let EndDate = this.SetTime(fnEventEndDate, eventEndDate);

                                    if (Startdate >= tempEventStartDate && Startdate <= tempEventEndDate) {
                                        itemArr.push({
                                            id: item.Id,
                                            recurrenceId: item.Id + ".0." + Startdate.toISOString().split('.')[0] + "Z",
                                            title: item[titleField],
                                            start: Startdate,
                                            end: EndDate,
                                            desc: item[descField] ? item[descField] : "",
                                            allDay: item[allDayEventField] ? item[allDayEventField] : false
                                        });
                                    }
                                }
                            }
                        }
                    }
                    //#endregion
                    //#region Yearly Events Process
                    else if (yearly) {
                        //Currently this is only for info
                        YearlyRecurrenceArr.push({
                            data: recurreceData
                        });

                        let attrYearly = xml.querySelector('yearly');
                        let attrYearlyByDay = xml.querySelector('yearlyByDay');
                        let yearFrequency = yearly.getAttribute('yearFrequency');

                        let tempEventStartDate = new Date(eventStartDate.getTime());
                        let tempEventEndDate = new Date(eventEndDate.getTime());

                        if (attrYearly) {

                            //#region XML Sample
                            /* <recurrence>
                                <rule>
                                    <firstDayOfWeek>su</firstDayOfWeek>
                                    <repeat>
                                        <yearly yearFrequency="1" month="5" day="5" />
                                    </repeat>
                                    <repeatForever>FALSE</repeatForever>
                                </rule>
                            </recurrence>
                            <recurrence>
                                <rule>
                                    <firstDayOfWeek>su</firstDayOfWeek>
                                    <repeat>
                                        <yearly yearFrequency="1" month="2" day="3" />
                                    </repeat>
                                    <repeatInstances>10</repeatInstances>
                                </rule>
                            </recurrence> */
                            //#endregion

                            let month = attrYearly.getAttribute('month');
                            let day = attrYearly.getAttribute('day');

                            for (let i = eventStartDate; i <= eventEndDate; i = new Date(new Date(new Date(i.setFullYear(i.getFullYear() + parseInt(yearFrequency))).setMonth(parseInt(month) - 1)).setDate(parseInt(day)))) {

                                let newTempStartDate = new Date(i.getTime());
                                let newTempEndDate = new Date(i.getTime());

                                newTempStartDate.setMonth(parseInt(month) - 1);
                                newTempStartDate.setDate(parseInt(day));
                                newTempEndDate.setMonth(parseInt(month) - 1);
                                newTempEndDate.setDate(parseInt(day));

                                if (newTempStartDate >= tempEventStartDate && newTempStartDate <= tempEventEndDate) {
                                    let Startdate = this.SetTime(newTempStartDate, eventStartDate);
                                    let EndDate = this.SetTime(newTempEndDate, eventEndDate);

                                    itemArr.push({
                                        id: item.Id,
                                        recurrenceId: item.Id + ".0." + Startdate.toISOString().split('.')[0] + "Z",
                                        title: item[titleField],
                                        start: Startdate,
                                        end: EndDate,
                                        desc: item[descField] ? item[descField] : "",
                                        allDay: item[allDayEventField] ? item[allDayEventField] : false
                                    });
                                }
                            }

                        }
                        else if (attrYearlyByDay) {

                            //#region XML Sample
                            /* <recurrence>
                                <rule>
                                    <firstDayOfWeek>su</firstDayOfWeek>
                                    <repeat>
                                        <yearlyByDay yearFrequency="1" su="TRUE" weekdayOfMonth="first" month="11" />
                                    </repeat>
                                    <repeatInstances>10</repeatInstances>
                                </rule>
                            </recurrence> */
                            //#endregion

                            let month = attrYearlyByDay.getAttribute('month');
                            let weekdayOfMonth = attrYearlyByDay.getAttribute('weekdayOfMonth');
                            let isSunday = attrYearlyByDay.getAttribute(sunday);
                            let isMonday = attrYearlyByDay.getAttribute(monday);
                            let isTuesday = attrYearlyByDay.getAttribute(tuesday);
                            let isWednesday = attrYearlyByDay.getAttribute(wednesday);
                            let isThrusday = attrYearlyByDay.getAttribute(thrusday);
                            let isFriday = attrYearlyByDay.getAttribute(friday);
                            let isSaturday = attrYearlyByDay.getAttribute(saturday);
                            let day = attrYearlyByDay.getAttribute("day");
                            let weekday = attrYearlyByDay.getAttribute("weekday");
                            let weekend_day = attrYearlyByDay.getAttribute("weekend_day");

                            let weekDayNumber = isSunday ? 0 : (
                                isMonday ? 1 : (
                                    isTuesday ? 2 : (
                                        isWednesday ? 3 : (
                                            isThrusday ? 4 : (
                                                isFriday ? 5 : (
                                                    isSaturday ? 6 : 0
                                                )
                                            )
                                        )
                                    )
                                )
                            );

                            for (let i = eventStartDate; i <= eventEndDate; i.setFullYear(i.getFullYear() + parseInt(yearFrequency))) {
                                let newTempStartDate = new Date(i.getTime());
                                let newTempEndDate = new Date(i.getTime());

                                newTempStartDate.setMonth(parseInt(month) - 1);

                                let fnEventStartDate;
                                let fnEventEndDate;

                                if (weekdayOfMonth === "last") {
                                    if (day) {
                                        fnEventStartDate = this.GetLastDay(newTempStartDate, DayWeekDayIndexArr, true, false, 0, false, false);
                                    }
                                    else if (weekday) {
                                        fnEventStartDate = this.GetLastDay(newTempStartDate, DayWeekDayIndexArr, false, false, 0, true, false);
                                    }
                                    else if (weekend_day) {
                                        fnEventStartDate = this.GetLastDay(newTempStartDate, DayWeekDayIndexArr, false, false, 0, false, true);
                                    }
                                    else if (isSunday || isMonday || isTuesday || isWednesday || isThrusday || isFriday || isSaturday) {
                                        fnEventStartDate = this.GetLastDay(newTempStartDate, DayWeekDayIndexArr, false, true, weekDayNumber, false, false);
                                    }
                                }
                                else if (weekdayOfMonth) {
                                    let apperanceNumber = weekdayOfMonth === "first" ? 1 : (
                                        weekdayOfMonth === "second" ? 2 : (
                                            weekdayOfMonth === "third" ? 3 : (
                                                weekdayOfMonth === "fourth" ? 4 : 0
                                            )
                                        )
                                    );

                                    if (day) {
                                        fnEventStartDate = this.GetRequiredDay(newTempStartDate, apperanceNumber, DayWeekDayIndexArr, true, false, 0, false, false);
                                    }
                                    else if (weekday) {
                                        fnEventStartDate = this.GetRequiredDay(newTempStartDate, apperanceNumber, DayWeekDayIndexArr, false, false, 0, true, false);
                                    }
                                    else if (weekend_day) {
                                        fnEventStartDate = this.GetRequiredDay(newTempStartDate, apperanceNumber, DayWeekDayIndexArr, false, false, 0, false, true);
                                    }
                                    else if (isSunday || isMonday || isTuesday || isWednesday || isThrusday || isFriday || isSaturday) {
                                        fnEventStartDate = this.GetRequiredDay(newTempStartDate, apperanceNumber, DayWeekDayIndexArr, false, true, weekDayNumber, false, false);
                                    }
                                }

                                if (fnEventStartDate && fnEventStartDate !== null) {
                                    fnEventEndDate = new Date(fnEventStartDate.getTime());

                                    let Startdate = this.SetTime(fnEventStartDate, eventStartDate);
                                    let EndDate = this.SetTime(fnEventEndDate, eventEndDate);

                                    if (Startdate >= tempEventStartDate && Startdate <= tempEventEndDate) {
                                        itemArr.push({
                                            id: item.Id,
                                            recurrenceId: item.Id + ".0." + Startdate.toISOString().split('.')[0] + "Z",
                                            title: item[titleField],
                                            start: Startdate,
                                            end: EndDate,
                                            desc: item[descField] ? item[descField] : "",
                                            allDay: item[allDayEventField] ? item[allDayEventField] : false
                                        });
                                    }
                                }
                            }
                        }
                    }
                    //#endregion
                    else {
                        console.log("Another Option Arrived");
                    }
                }
            }
            catch (ex) {
            }
            resolve(itemArr);
        });
    }

    private EditRecurringEvents(itemArr: ICalendarEvents[], editEventsArr: any, titleField: string, descField: string, startDateField: string, endDateField: string): any {
        for (let i = 0; i < editEventsArr.length; i++) {
            if (editEventsArr.length > 1) {
                itemArr.filter((item, index) => {
                    if (item["recurrenceId"].toString() === editEventsArr[i]["MasterSeriesItemID"] + ".0." + editEventsArr[i]["RecurrenceID"].toString()) {
                        itemArr[index]["recurrenceId"] = editEventsArr[i]["Id"] + ".1." + itemArr[index]["id"];
                        itemArr[index]["title"] = editEventsArr[i][titleField];
                        itemArr[index]["desc"] = editEventsArr[i][descField] ? editEventsArr[i][descField] : "";
                        itemArr[index]["start"] = new Date(editEventsArr[i][startDateField]);
                        itemArr[index]["end"] = new Date(editEventsArr[i][endDateField]);
                        return;
                    }
                });
            }
        }
        return itemArr;
    }

    private RemoveDeletedEvents(itemArr: ICalendarEvents[], deletedEventsArr: any): any {
        for (let i = 0; i < deletedEventsArr.length; i++) {
            if (deletedEventsArr[i]["RecurrenceID"]) {
                itemArr.filter((item, index) => {
                    if (item["recurrenceId"].toString() === deletedEventsArr[i]["MasterSeriesItemID"] + ".0." + deletedEventsArr[i]["RecurrenceID"].toString()) {
                        delete itemArr[index];
                    }
                });
            }
        }
        var itemArr = itemArr.filter((x) => {
            return x !== undefined;
        });
        return itemArr;
    }
}