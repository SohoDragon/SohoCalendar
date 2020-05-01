export default interface ICalendarEvents {
  id: number;
  title: string;
  start: Date;
  end: Date;
  desc: string;
  allDay: boolean;
}