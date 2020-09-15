using Ical.Net;
using Ical.Net.CalendarComponents;
using Ical.Net.DataTypes;
using Ical.Net.Serialization;
using MailKit.Net.Smtp;
using MimeKit;
using System;

namespace EmailSender
{
    public static class EmailSenderMailKit
    {
        public static void SendMessage(MimeMessage message)
        {
            using (var client = new SmtpClient())
            {
                client.ServerCertificateValidationCallback = (s, c, h, e) => true;
                client.Connect("smtp.office365.com", 587, false);
                client.AuthenticationMechanisms.Remove("XOAUTH2");
                client.Authenticate("username", "password");
                client.Send(message);
                client.Disconnect(true);
            }
        }

        private static string CreateCalendarEntry(DateTime? start, DateTime? end, string title, string description, string location)
        {
            Calendar iCal = new Calendar();
            iCal.Method = "PUBLISH";
            // Create the event, and add it to the iCalendar
            CalendarEvent evt = iCal.Create<CalendarEvent>();
            // Set information about the event
            evt.Start = new CalDateTime(start.Value);
            evt.End = new CalDateTime(end.Value); // This also sets the duration  
            evt.Description = description;
            evt.Location = location;
            evt.Summary = title;
            // Create a reminder 24h before the event
            Alarm reminder = new Alarm();
            reminder.Action = AlarmAction.Display;
            reminder.Trigger = new Trigger(new TimeSpan(-24, 0, 0));
            evt.Alarms.Add(reminder);

            var serializer = new CalendarSerializer();
            var serializedCalendar = serializer.SerializeToString(iCal);

            //return iCal;
            return serializedCalendar;
        }
    }
}