using OfficeOpenXml;
using System;
using System.IO;
using System.Text;

namespace Classes_Calendar
{
    ///<summary>
    /// Klasse welche die Appointment-Liste in eine Ical-Datei umwandelt
    ///</summary>
    public class IcalConverter
    {   
        
        
        public static void ConvertToIcalFile(List<Appointment> appointments, string icalFilePath)
        {
            using (StreamWriter writer = new StreamWriter(icalFilePath))
            {
                writer.WriteLine("BEGIN:VCALENDAR");
                writer.WriteLine("VERSION:2.0");
                writer.WriteLine("PRODID:-//BoschVM-Tyco//BetterPlan//DE");

                foreach (var appointment in appointments)
                {
                    writer.WriteLine("BEGIN:VEVENT");
                    writer.WriteLine($"SUMMARY:{appointment.AppointmentName}");
                    writer.WriteLine($"LOCATION:{appointment.Location}");
                    writer.WriteLine($"DTSTART:{appointment.StartTime:yyyyMMddTHHmmss}");
                    writer.WriteLine($"DTEND:{appointment.EndTime:yyyyMMddTHHmmss}");
                    writer.WriteLine("END:VEVENT");
                }

                writer.WriteLine("END:VCALENDAR");
            }
        }
    }
}
