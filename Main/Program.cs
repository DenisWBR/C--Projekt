using System.Collections.Generic;
using System;
using System.IO;
using Classes_Calendar;

    class Program
    {
        static void Main()
        {
            string excelFilePath = @"C:\Users\49176\Desktop\Dokumente_Studium\C#\Exceldateien_Projekt\VP_20224.xlsx";

            ExcelReader excelReader = new ExcelReader(excelFilePath);

            //IcalConverter.ConvertToIcalFile(appointments, icalFilePath);

            List<Course> courses = excelReader.ReadCourses();

            int courseRow = excelReader.ReadCoursesRow();

            List<TimeSlot> timeSlots = excelReader.ReadTimes();

            //Excel Reader           

            foreach (Course course in courses)
            {
                Console.WriteLine($"Kursname: {course.CourseName}, Kursraum: {course.CourseRoom}, Excel-Spalte: {course.CourseColumnIndex}");
            }

            System.Console.WriteLine("Welchen Kurs möchten sie konvertieren?(Spaltennummer eingeben)");
            string? userInput = Console.ReadLine();
            int courseColumn = Convert.ToInt32(userInput);

            List<Subject> subjects = excelReader.ReadSubjects(courseColumn, courseRow);

            foreach (Subject subject in subjects)
            {
                Console.WriteLine($"Vorlesung: {subject.SubjectName}, Vorlesungsort: {subject.Location}, Excel-Reihe: {subject.SubjectRowIndex}");
            }

            foreach (TimeSlot timeSlot in timeSlots)
            {
                Console.WriteLine($"Beginn: {timeSlot.StartTime}, Ende: {timeSlot.EndTime}, Excel-Reihe: {timeSlot.TimeSlotRowIndex}");
            }

            //Appointment Converter

            var groupedSubjects = AppointmentSorter.GroupSubjects(subjects);

            foreach (var subject in groupedSubjects)
            {
                Console.WriteLine($"Vorlesung: {subject.SubjectName}, Vorlesungsort: {subject.Location}");

                Console.WriteLine($"Block: {subject.SubjectRowIndexStart} - {subject.SubjectRowIndexEnd}");
            }
            var appointments = AppointmentTimeAssigner.AssignAppointmentTimes(subjects,timeSlots);

            foreach (var appointment in appointments)
            {
                Console.WriteLine($"Vorlesung: {appointment.AppointmentName}, Vorlesungsort: {appointment.Location}");

                Console.WriteLine($"Startzeit: {appointment.StartTime}, Endzeit: {appointment.EndTime}");   
            }

            //Ical Converter
            //string icalFilePath = @"C:\Users\49176\Desktop\Dokumente_Studium\C#\Exceldateien_Projekt";
            //IcalConverter.ConvertToIcalFile(appointments, "Kaledner.ics");
            //So wie es jetzt dransteht wird die Kalenderdatei im selben ordner wie program.cs erstellt
            //es müsste aber so aussehen:
        
            string directory = @"C:\Users\49176\Desktop\Dokumente_Studium\C#\Exceldateien_Projekt";
            string fileName = "neu.ics";
            string icalFilePath = Path.Combine(directory, fileName);
            IcalConverter.ConvertToIcalFile(appointments, icalFilePath);

        }  
    }

