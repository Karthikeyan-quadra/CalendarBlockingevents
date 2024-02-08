// import * as React from 'react';
// import FullCalendar from '@fullcalendar/react';
// import dayGridPlugin from '@fullcalendar/daygrid';
// import timeGridPlugin from '@fullcalendar/timegrid';
// import interactionPlugin from '@fullcalendar/interaction';
// import { useEffect, useState } from "react";
// import { MSGraphClientV3 } from "@microsoft/sp-http";
// import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";

// import { ICalendarProps } from './ICalendarProps';

// // Add your custom styles here

// export default function Calendar(props: ICalendarProps) {
//   const [events, setEvents] = useState<MicrosoftGraph.Event[]>([]);

//   useEffect(() => {
//     props.context.msGraphClientFactory
//       .getClient("3")
//       .then((client: MSGraphClientV3) => {
//         client
//           .api('/me/calendar/events')
//           .version("v1.0")
//           .select("*")
//           .get((error: any, eventsResponse, rawResponse?: any) => {
//             if (error) {
//               console.error("Message is: " + error);
//               return;
//             }

//             const Events: MicrosoftGraph.Event[] = eventsResponse.value;
//             setEvents(Events);
//           });
//       });
//   }, [props.context.msGraphClientFactory]);

//   return (
//     <>
//       {/* <style>{customStyles}</style> */}
//       <FullCalendar
//         plugins={[dayGridPlugin, timeGridPlugin, interactionPlugin]}
//         initialView="dayGridMonth"
//         headerToolbar={{
//           left: 'prev,next today',
//           center: 'title',
//           right: 'dayGridMonth,timeGridWeek,timeGridDay'
//         }}
//         events={events.map(event => ({
//           title: event.subject ?? '',
//           start: event.start?.dateTime ?? '',
//           end: event.end?.dateTime ?? '',
//         }))}
//         eventContent={(eventContent) => {
//           return (
//             <>
//               <div style={{ width: '100%', textAlign: 'center', backgroundColor: 'rgba(169, 211, 242, 0.9)', color: 'black', borderRadius: '5px', padding: '2px', fontSize: '14px', borderLeft: '5px solid rgba(0, 120, 212, 0.9)', height:'22px', fontWeight:'400', overflow: 'hidden', whiteSpace: 'nowrap',
//             }}>
//                 {eventContent.event.start?.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' })} - {eventContent.event.title}
//               </div>
//             </>
//           );
//         }}
//       />
//     </>
//   );
// }


// import * as React from "react";
// import { useEffect, useState } from "react";
// import { MSGraphClientV3 } from "@microsoft/sp-http";
// import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
// import FullCalendar from "@fullcalendar/react";
// import dayGridPlugin from "@fullcalendar/daygrid";
// import timeGridPlugin from "@fullcalendar/timegrid";
// import interactionPlugin from "@fullcalendar/interaction";
// import { ICalendarProps } from "./ICalendarProps";
// import styles from "./Calendar.module.scss";
// import OverlayTrigger from "react-bootstrap/OverlayTrigger";
// import Popover from "react-bootstrap/Popover";

// import * as React from 'react';
// import { useEffect, useState } from "react";
// import { MSGraphClientV3 } from "@microsoft/sp-http";
// import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
// import FullCalendar from '@fullcalendar/react';
// import dayGridPlugin from '@fullcalendar/daygrid';
// import timeGridPlugin from '@fullcalendar/timegrid';
// import interactionPlugin from '@fullcalendar/interaction';
// import { ICalendarProps } from './ICalendarProps';
// import styles from './Calendar.module.scss';
// import OverlayTrigger from 'react-bootstrap/OverlayTrigger';
// import Popover from 'react-bootstrap/Popover';
// import 'bootstrap/dist/css/bootstrap.min.css';


// interface IFormattedEvent {
//   subject: string;
//   startDate: string;
//   endDate: string;
//   startTime: string;
//   endTime: string;
//   eventDate: string;
//   bodyPreview?: string;
//   joinUrl?: string;
// }

// var customStyles = `

//     a {
//       color: #000;
//       text-decoration: none;

//     }
//     .fc .fc-button-primary:disabled {
//       background-color: #7787A9;
//       border-color: #7787A9;
//       opacity: 1;
//     }

//     .fc .fc-button-primary {
//       background-color: #293859;
//       border-color: #293859;
//     }

//     .fc .fc-button-primary:not(:disabled).fc-button-active {
//       back
//     }

//     :root {
//       --fc-today-bg-color: #ececec;
//       --fc-event-bg-color: #91afd9db;
//     --fc-event-border-color: #91afd9db;
//     }

//     .popover {
//       max-width: none !important;
//       /* Ensure the popover does not have a max-width */
//     }
//     .popover-arrow {
//       border-right-color: #fff !important;
//     }
//   `;

// const EventsCalendar: React.FC<ICalendarProps> = (props: any) => {
//   const [events, setEvents] = useState<MicrosoftGraph.Event[]>([]);

//   useEffect(() => {
//     props.context.msGraphClientFactory
//       .getClient("3")
//       .then((client: MSGraphClientV3) => {
//         client
//           .api("me/calendar/events")
//           .version("v1.0")
//           .select("*")
//           .get((error: any, eventsResponse, rawResponse?: any) => {
//             if (error) {
//               console.error("Message is: " + error);
//               return;
//             }

//             const calendarEvents: MicrosoftGraph.Event[] = eventsResponse.value;
//             setEvents(
//               calendarEvents.map((event) => ({
//                 ...event,
//                 joinUrl: event.onlineMeeting?.joinUrl || "",
//                 bodyPreview: event.bodyPreview || "", // Use bodyPreview if available, otherwise default to an empty string
//               }))
//             );

//             console.log("CalendarEvents", calendarEvents);
//           });
//       });
//   }, [props.context.msGraphClientFactory]);

//   const eventContent = (eventInfo: any) => {
//     const formattedEvent: IFormattedEvent = {
//       subject: eventInfo.event.title,
//       startDate: eventInfo.event.startStr,
//       startTime: eventInfo.event.start.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' }).toUpperCase(),
//       endDate: eventInfo.event.endStr,
//       endTime: eventInfo.event.end.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' }).toUpperCase(),
//       eventDate: eventInfo.event.start.toString(),
//       bodyPreview: eventInfo.event.extendedProps.bodyPreview,
//       joinUrl: eventInfo.event.extendedProps.joinUrl,
//     };
//     console.log("Join URL:", formattedEvent.joinUrl);
//     console.log("Formatted Event", formattedEvent);
//     console.log("EventInfo", eventInfo);
//     console.log("Event Content", eventContent);

//     const popover = (
//       <Popover
//         id={`popover-${formattedEvent.startDate}`}
//         className={styles.popoverBox}
//       >
//         <Popover.Header as="h3" className={styles.popheader}>
//           <b>Calendar</b> - <span>{props.context.pageContext.user.email}</span>
//         </Popover.Header>
//         <Popover.Body>
//           <div className={styles.popBody}>
//             <div className={styles.popheading}>
//               <img
//                 src={require("../assets/Icon1.svg")}
//                 alt="Icon"
//                 className={styles.popoverIcon}
//               />
//               <span className={styles.contentStyle}>
//                 {formattedEvent.subject}
//               </span>
//             </div>
//             <div className={styles.popContent}>
//               <img
//                 src={require("../assets/Icon2.svg")}
//                 alt="Icon"
//                 className={styles.popoverIcon}
//               />
//               <span className={styles.contentStyle}>
//                 {`${formattedEvent.eventDate.substring(
//                   0,
//                   3
//                 )}, ${formattedEvent.eventDate.substring(
//                   4,
//                   10
//                 )} ${formattedEvent.startTime}
//                  - ${formattedEvent.endTime}`}
//               </span>
//             </div>
//             <div
//               className={styles.popContent}
//               style={{ display: formattedEvent.bodyPreview ? "flex" : "none" }}
//             >
//               <img
//                 src={require("../assets/Icon3.svg")}
//                 alt="Icon"
//                 className={styles.popoverIcon}
//               />
//               <p className={styles.contentStyle}>
//                 {formattedEvent.bodyPreview}
//               </p>
//             </div>
//             <div style={{ display: formattedEvent.joinUrl ? "flex" : "none" }}>
//               <button className={styles.joinBtn}>
//                 <a href={formattedEvent.joinUrl} target="_blank">
//                   Join
//                 </a>
//               </button>
//             </div>
//           </div>
//         </Popover.Body>
//       </Popover>
//     );
//     return (
//       <OverlayTrigger
//         trigger="click"
//         placement="right"
//         overlay={popover}
//         rootClose={true}
//       >
//         <button className={styles.popoverButton}>
//           <span>{formattedEvent.startTime} </span>
//           <b> {formattedEvent.subject}</b>
//         </button>
//       </OverlayTrigger>
//     );
//   };

//   return (
//     <div className={styles.calendarApp}>
//       <style>{customStyles}</style>
//       <div className={styles.calendarAppMain}>
//         <FullCalendar
//           plugins={[dayGridPlugin, timeGridPlugin, interactionPlugin]}
//           headerToolbar={{
//             left: "prev,next today",
//             center: "title",
//             right: "dayGridMonth,timeGridWeek,timeGridDay",
//           }}
//           initialView="dayGridMonth"
//           customButtons={{
//             customPrev: { text: "Prev" },
//             customNext: { text: "Next" },
//             customToday: { text: "Today" },
//           }}
//           buttonText={{
//             prev: "<",
//             next: ">",
//             today: "Today",
//             dayGridMonth: "Month",
//             timeGridWeek: "Week",
//             timeGridDay: "Day",
//           }}
//           events={events.map((event: any) => ({
//             title: event.subject,
//             start: event.start.dateTime,
//             end: event.end.dateTime,
//             bodyPreview: event.bodyPreview,
//             joinUrl: event.joinUrl,
//           }))}
//           eventContent={eventContent}
//         />
//       </div>
//     </div>
//   );
// };

// export default EventsCalendar;


import * as React from 'react';
import { useEffect, useState } from "react";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
import FullCalendar from '@fullcalendar/react';
import dayGridPlugin from '@fullcalendar/daygrid';
import timeGridPlugin from '@fullcalendar/timegrid';
import interactionPlugin from '@fullcalendar/interaction';
import { ICalendarProps } from './ICalendarProps';
import styles from './Calendar.module.scss';
import OverlayTrigger from 'react-bootstrap/OverlayTrigger';
import Popover from 'react-bootstrap/Popover';
import 'bootstrap/dist/css/bootstrap.min.css';

interface IFormattedEvent {
  subject: string;
  startDate: string;
  endDate: string;
  startTime: string;
  endTime: string;
  eventDate: string;
  bodyPreview?: string;
  joinUrl?: string;
}

var customStyles = `
  a {
    color: #000;
    text-decoration: none;
  }
  .fc .fc-button-primary:disabled {
    background-color: #7787A9;
    border-color: #7787A9;
    opacity: 1;
  }
  .fc .fc-button-primary {
    background-color: #293859;
    border-color: #293859;
  }
  .fc .fc-button-primary:not(:disabled).fc-button-active {
    back
  }
  :root {
    --fc-today-bg-color: #ececec;
    --fc-event-bg-color: #91afd9db;
    --fc-event-border-color: #91afd9db;
  }
  .popover {
    max-width: none !important;
    /* Ensure the popover does not have a max-width */
  }
  .popover-arrow {
    border-right-color: #fff !important;
  }
  `;

const EventsCalendar: React.FC<ICalendarProps> = (props: any) => {
  const [events, setEvents] = useState<MicrosoftGraph.Event[]>([]);

  useEffect(() => {
    props.context.msGraphClientFactory
      .getClient("3")
      .then((client: MSGraphClientV3) => {
        client
          .api("me/calendar/events")
          .version("v1.0")
          .select("*")
          .get((error: any, eventsResponse, rawResponse?: any) => {
            if (error) {
              console.error("Message is: " + error);
              return;
            }

            const calendarEvents: MicrosoftGraph.Event[] = eventsResponse.value;
            setEvents(
              calendarEvents.map((event) => ({
                ...event,
                joinUrl: event.onlineMeeting?.joinUrl || "",
                bodyPreview: event.bodyPreview || "", // Use bodyPreview if available, otherwise default to an empty string
              }))
            );

            console.log("CalendarEvents", calendarEvents);
          });
      });
  }, [props.context.msGraphClientFactory]);

  const eventContent = (eventInfo: any) => {
    const formattedEvent: IFormattedEvent = {
      subject: eventInfo.event.title,
      startDate: eventInfo.event.startStr,
      startTime: eventInfo.event.start.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' }).toUpperCase(),
      endDate: eventInfo.event.endStr,
      endTime: eventInfo.event.end.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' }).toUpperCase(),
      eventDate: eventInfo.event.start.toString(),
      bodyPreview: eventInfo.event.extendedProps.bodyPreview,
      joinUrl: eventInfo.event.extendedProps.joinUrl,
    };

    const isEventOver = new Date() > new Date(formattedEvent.endDate);

    const popover = (
      <Popover
        id={`popover-${formattedEvent.startDate}`}
        className={styles.popoverBox}
      >
        <Popover.Header as="h3" className={styles.popheader}>
          <b>Calendar</b> - <span>{props.context.pageContext.user.email}</span>
        </Popover.Header>
        <Popover.Body>
          <div className={styles.popBody}>
            <div className={styles.popheading}>
              <img
                src={require("../assets/Icon1.svg")}
                alt="Icon"
                className={styles.popoverIcon}
              />
              <span className={styles.contentStyle}>
                {formattedEvent.subject}
              </span>
            </div>
            <div className={styles.popContent}>
              <img
                src={require("../assets/Icon2.svg")}
                alt="Icon"
                className={styles.popoverIcon}
              />
              <span className={styles.contentStyle}>
                {`${formattedEvent.eventDate.substring(
                  0,
                  3
                )}, ${formattedEvent.eventDate.substring(
                  4,
                  10
                )} ${formattedEvent.startTime}
                 - ${formattedEvent.endTime}`}
              </span>
            </div>
            <div
              className={styles.popContent}
              style={{ display: formattedEvent.bodyPreview ? "flex" : "none" }}
            >
              <img
                src={require("../assets/Icon3.svg")}
                alt="Icon"
                className={styles.popoverIcon}
              />
              <p className={styles.contentStyle}>
                {formattedEvent.bodyPreview}
              </p>
            </div>
            <div style={{ display: formattedEvent.joinUrl && !isEventOver ? "flex" : "none" }}>
              <button className={styles.joinBtn}>
                <a href={formattedEvent.joinUrl} target="_blank">
                  Join
                </a>
              </button>
            </div>
          </div>
        </Popover.Body>
      </Popover>
    );

    return (
      <OverlayTrigger
        trigger="click"
        placement="right"
        overlay={popover}
        rootClose={true}
      >
        <button className={styles.popoverButton}>
          <span>{formattedEvent.startTime} </span>
          <b> {formattedEvent.subject}</b>
        </button>
      </OverlayTrigger>
    );
  };

  return (
    <div className={styles.calendarApp}>
      <style>{customStyles}</style>
      <div className={styles.calendarAppMain}>
        <FullCalendar
          plugins={[dayGridPlugin, timeGridPlugin, interactionPlugin]}
          headerToolbar={{
            left: "prev,next today",
            center: "title",
            right: "dayGridMonth,timeGridWeek,timeGridDay",
          }}
          initialView="dayGridMonth"
          customButtons={{
            customPrev: { text: "Prev" },
            customNext: { text: "Next" },
            customToday: { text: "Today" },
          }}
          buttonText={{
            prev: "<",
            next: ">",
            today: "Today",
            dayGridMonth: "Month",
            timeGridWeek: "Week",
            timeGridDay: "Day",
          }}
          events={events.map((event: any) => ({
            title: event.subject,
            start: event.start.dateTime,
            end: event.end.dateTime,
            bodyPreview: event.bodyPreview,
            joinUrl: event.joinUrl,
          }))}
          eventContent={eventContent}
        />
      </div>
    </div>
  );
};

export default EventsCalendar;
