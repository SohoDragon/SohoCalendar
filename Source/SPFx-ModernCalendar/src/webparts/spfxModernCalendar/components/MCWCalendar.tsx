import * as React from 'react';
import { IMCWCalendarProps } from '../models/IMCWCalendarProps';
import { IMCWCalendarState } from "../models/IMCWCalendarState";
import styles from "../components/cutomStyle.module.scss";
import calendarstyles from "../components/GraphCalendar.module.scss";

import FullCalendar from '@fullcalendar/react';
import dayGridPlugin from '@fullcalendar/daygrid';
import listPlugin from '@fullcalendar/list';
import timeGridPlugin from '@fullcalendar/timegrid';
import momentPlugin from '@fullcalendar/moment';

import '@fullcalendar/core/main.css';
import '@fullcalendar/daygrid/main.css';
import '@fullcalendar/list/main.css';
import '@fullcalendar/timegrid/main.css';

import defaultEvents from "../data/defaultEvents";

export default class MCWCalendar extends React.Component<IMCWCalendarProps, IMCWCalendarState> {
    constructor(props: IMCWCalendarProps, state: IMCWCalendarState) {
        super(props);

        this.state = {
            showDialog: false,
            title: "",
            startDate: "",
            endDate: "",
            description: ""
        };
    }

    private _OpenModal = (e): void => {
        e = e.event;
        let desc = e._def.extendedProps.desc;
        this.setState({
            showDialog: true,
            title: e.title,
            startDate: e.start.toString(),
            endDate: e.end.toString(),
            description: desc
        });
    }

    private _CloseModal = (): any => {
        this.setState({
            showDialog: false
        });
    }

    public render(): React.ReactElement<IMCWCalendarProps> {
        return (
            <div className={calendarstyles.graphCalendar}>
                <div className={this.state.showDialog == true ? styles.model : styles.hide}>
                    <div className={styles.spfxModernCalendar}>
                        <span style={
                            {
                                background: "radial-gradient(" + this.props.EventBGColor_compo + ", transparent)",
                                color: this.props.EventTitleColor_compo
                            }
                        } className={styles.popupHeader}>{this.state.title}</span>

                        <div className={styles.row}>
                            <div className={styles.column}>
                                <div className={styles.eventTimeStyle}>
                                    <div>
                                        <span style={
                                            {
                                                background: "radial-gradient(" + this.props.EventBGColor_compo + ", transparent)",
                                                color: this.props.EventTitleColor_compo
                                            }
                                        } className={styles.popupContent}>Event Time</span>
                                    </div>
                                    <div>
                                        <span>{(new Date(this.state.startDate)).toLocaleString()} - {(new Date(this.state.endDate)).toLocaleString()}</span>
                                    </div>
                                </div>
                                <div className={styles.eventTimeStyle}>
                                    <div>
                                        <span style={
                                            {
                                                background: "radial-gradient(" + this.props.EventBGColor_compo + ", transparent)",
                                                color: this.props.EventTitleColor_compo
                                            }
                                        } className={styles.popupContent}>Event Description</span>
                                    </div>
                                    <div>
                                        <span dangerouslySetInnerHTML={{ __html: this.state.description }}></span>
                                    </div>
                                </div>
                                <div>
                                    <a style={
                                        {
                                            background: this.props.EventBGColor_compo,
                                            color: this.props.EventTitleColor_compo
                                        }
                                    } onClick={this._CloseModal} className={styles.button}>
                                        <span className={styles.label}>Close</span>
                                    </a>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
                <div><h1>{this.props.WebpartTitle_compo}</h1></div>
                <FullCalendar
                    header={{
                        left: 'prev,next, today',
                        center: 'title',
                        right: 'dayGridMonth,dayGridWeek,dayGrid,listWeek,timeGrid'
                    }}
                    defaultView="dayGridMonth"
                    plugins={[dayGridPlugin, listPlugin, timeGridPlugin, momentPlugin]}
                    aspectRatio={2}
                    eventBackgroundColor={this.props.EventBGColor_compo}
                    eventTextColor={this.props.EventTitleColor_compo}
                    height="auto"
                    eventLimit={true}
                    eventClick={(events) => {
                        this._OpenModal(events);
                    }
                    }
                    events={this.props.Events} />
            </div>

        );
    }
}