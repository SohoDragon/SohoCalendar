export default interface ICalendarEvents {
  id: string;
  recurrenceId: string;
  title: string;
  start: Date;
  end: Date;
  desc: string;
  allDay: boolean;
}