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

import classNames from 'classnames';

export default class MCWCalendar extends React.Component<IMCWCalendarProps, IMCWCalendarState> {
    constructor(props: IMCWCalendarProps, state: IMCWCalendarState) {
        super(props);

        this.state = {
            showDialog: false,
            title: "",
            startDate: "",
            endDate: "",
            description: "",
            viewLink: ""
        };
    }

    private _OpenModal = (e): void => {
        e = e.event;
        let desc = e._def.extendedProps.desc;
        let viewEventLink = "";

        if (this.props.DisplayFormURL_combo) {
            let sourceURL = window.location.href;
            let webURL = this.props.context.pageContext.web.absoluteUrl;
            let itemId = e._def.extendedProps.recurrenceId ? e._def.extendedProps.recurrenceId : e.id;

            if (this.props.DisplayFormURL_combo.indexOf(webURL) < 0) {
                viewEventLink = webURL;

                if (this.props.DisplayFormURL_combo.indexOf("/") !== 0) {
                    viewEventLink += "/";
                }
            }

            viewEventLink += this.props.DisplayFormURL_combo + "?ID=" + itemId + "&Source=" + sourceURL;
        }
        this.setState({
            showDialog: true,
            title: e.title,
            startDate: e.start.toString(),
            endDate: e.end.toString(),
            description: desc,
            viewLink: viewEventLink
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
                                <div className={styles.buttonGroup}>
                                    <a style={
                                        {
                                            background: this.props.EventBGColor_compo,
                                            color: this.props.EventTitleColor_compo
                                        }
                                    } href={this.state.viewLink} className={this.state.viewLink ? classNames({ [styles.button]: true, [styles.ViewLinkButton]: true }) : styles.hide}>
                                        <span className={styles.label}>View Event</span>
                                    </a>
                                    <a style={
                                        {
                                            background: this.props.EventBGColor_compo,
                                            color: this.props.EventTitleColor_compo
                                        }
                                    } onClick={this._CloseModal} className={this.state.viewLink ? styles.button : classNames({ [styles.button]: true, [styles.onlyCancel]: true })}>
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
                    views={ {
                        dayGridMonth: {
                          eventLimit: 4
                        },
                        dayGridWeek: {
                            eventLimit: 4
                        },
                        dayGrid: {
                            eventLimit: 4
                        }
                      }
                    }
                    defaultView="dayGridMonth"
                    plugins={[dayGridPlugin, listPlugin, timeGridPlugin, momentPlugin]}
                    aspectRatio={5}
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