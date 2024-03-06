using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Linq;

namespace Classes_Calendar
{
    public class AppointmentSorter
    {
        public static List<Subject> GroupSubjects(List<Subject> subjects)
        {
            var groupedSubjects = new List<Subject>();
            var sortedSubjects = subjects.OrderBy(v => v.SubjectName).ThenBy(v => v.Location).ThenBy(v => v.SubjectRowIndex).ToList();
            Subject? currentGroup = null;
            foreach (var subject in sortedSubjects)
            {
                if (currentGroup == null || currentGroup.SubjectName != subject.SubjectName || currentGroup.Location != subject.Location || currentGroup.SubjectRowIndexEnd + 1 != subject.SubjectRowIndex)
                {
                    if (currentGroup != null)
                    {
                        groupedSubjects.Add(currentGroup);
                    }
                    currentGroup = new Subject(subject.SubjectName, subject.Location, subject.SubjectRowIndex);
                }
                else
                {
                    currentGroup.SubjectRowIndexEnd = subject.SubjectRowIndex;
                }

            }
            if (currentGroup != null)
            {
                groupedSubjects.Add(currentGroup);
            }
            return groupedSubjects;
        }
    }

    public class AppointmentTimeAssigner
    {
        public static List<Appointment> AssignAppointmentTimes(List<Subject> subjects, List<TimeSlot> timeSlots)
        {
            List<Appointment> appointments = new List<Appointment>();

            var groupedSubjects = AppointmentSorter.GroupSubjects(subjects);


            foreach (var subject in groupedSubjects)
            {

                DateTime? startTime = timeSlots.FirstOrDefault(ts => ts.TimeSlotRowIndex == subject.SubjectRowIndexStart)?.StartTime;
                DateTime? endTime = timeSlots.FirstOrDefault(ts => ts.TimeSlotRowIndex == subject.SubjectRowIndexEnd)?.EndTime;

                if (startTime != null && endTime != null)
                {
                    Appointment appointment = new Appointment(subject.SubjectName, subject.Location, startTime.Value, endTime.Value);
                    appointments.Add(appointment);
                }

            }

            return appointments;
        }
    }


}
