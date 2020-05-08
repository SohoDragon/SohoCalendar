import * as React from 'react';
import { IMCWCalendarProps } from '../models/IMCWCalendarProps';
import { Calendar, momentLocalizer } from 'react-big-calendar';
import "react-big-calendar/lib/css/react-big-calendar.css";
import * as moment from "moment";
import { IMCWCalendarState } from "../models/IMCWCalendarState";
import styles from "../components/cutomStyle.module.scss";

const localizer = momentLocalizer(moment);

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
        this.setState({
            showDialog: true,
            title: e.title,
            startDate: e.start.toString(),
            endDate: e.end.toString(),
            description: e.desc
        });
    }

    private _CloseModal = (): any => {
        this.setState({
            showDialog: false
        });
    }

    private eventStyleGetter = (event, start, end, isSelected): any => {
        var style = {
            backgroundColor: this.props.EventBGColor_compo,
            color: this.props.EventTitleColor_compo,
            display: 'block'
        };
        return {
            style: style
        };
    }

    public render(): React.ReactElement<IMCWCalendarProps> {
        return (
            <div>
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
                                        <span className={styles.popupContent}>Event Time</span>
                                    </div>
                                    <div>
                                        <span>{(new Date(this.state.startDate)).toLocaleString()} - {(new Date(this.state.endDate)).toLocaleString()}</span>
                                    </div>
                                </div>
                                <div className={styles.eventTimeStyle}>
                                    <div>
                                        <span className={styles.popupContent}>Event Description</span>
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
                <Calendar
                    localizer={localizer}
                    events={this.props.Events}
                    startAccessor="start"
                    endAccessor="end"
                    style={{ height: 500 }}
                    popup={true}
                    eventPropGetter={(this.eventStyleGetter)}
                    onSelectEvent={(events) => {
                        this._OpenModal(events);
                    }
                    }
                />
            </div>

        );
    }
}