using Ical.Net;
using Ical.Net.CalendarComponents;
using Ical.Net.DataTypes;
using Ical.Net.Serialization;
using System;
using System.Net.Mail;

namespace EmailSender
{
    public static class EmailSenderNetMail
    {
        public static void SendCalendarEvent(Calendar calendar)
        {
            CalendarSerializer serializer = new CalendarSerializer(new SerializationContext());

            //preparing message
            MailMessage msg = new MailMessage();
            msg.From = new MailAddress("testingmordo@outlook.com", "Credit Inquiry");
            msg.To.Add(new MailAddress("pietrusjakub@outlook.com", "Song John"));
            msg.To.Add(new MailAddress("jpieetrus@gmail.com", "jpi"));
            msg.Subject = "CS Inquiry";

            System.Net.Mime.ContentType ct = new System.Net.Mime.ContentType("text/calendar");
            ct.Parameters.Add("method", "REQUEST");
            ct.Parameters.Add("name", "meeting.ics");
            AlternateView avCal = AlternateView.CreateAlternateViewFromString(serializer.SerializeToString(calendar), ct);
            msg.AlternateViews.Add(avCal);



            // sending
            SmtpClient sc = new SmtpClient("smtp.office365.com", 587);
            sc.UseDefaultCredentials = false;
            sc.Credentials = new System.Net.NetworkCredential("username", "password");
            sc.DeliveryMethod = SmtpDeliveryMethod.Network;
            sc.EnableSsl = true;

            sc.Send(msg);
        }

        public static Calendar CreateCalendar()
        {
            DateTime dtStart = new DateTime(2020, 09, 15);
            DateTime dtEnd = dtStart.AddDays(2);

            CalendarEvent e = new CalendarEvent()
            {
                DtStart = new CalDateTime(dtStart),
                DtEnd = new CalDateTime(dtEnd),
                DtStamp = new CalDateTime(DateTime.Now),
                IsAllDay = true,
                Sequence = 0,
                Transparency = TransparencyType.Transparent,
                Description = "Test with iCal.Net",
                Priority = 5,
                Class = "PUBLIC",
                Location = "New York",
                Summary = "Tested with iCal.Net Summary",
                Uid = Guid.NewGuid().ToString(),
                Organizer = new Organizer()
                {
                    CommonName = "John, Song",
                    Value = new Uri("mailto:testingmordo@outlook.com")
                },
            };

            e.Attendees.Add(new Attendee()
            {
                CommonName = "John, Song",
                ParticipationStatus = "REQ-PARTICIPANT",
                Rsvp = true,
                Value = new Uri("mailto:pietrusjakub@outlook.com")
            });


            Alarm alarm = new Alarm()
            {
                Action = AlarmAction.Display,
                Trigger = new Trigger(TimeSpan.FromDays(-1)),
                Summary = "Inquiry due in 1 day"
            };

            e.Alarms.Add(alarm);

            Calendar c = new Calendar();

            c.Events.Add(e);

            c.Method = "REQUEST";

            return c;
        }
    }
}