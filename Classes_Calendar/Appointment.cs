namespace Classes_Calendar

{
public class Appointment
{
    //Eigenschaften
    public string AppointmentName {get; set;}

    public object Location{get; set;}

    public DateTime StartTime{get; set;}

    public DateTime EndTime{get; set; }

    //Konstruktor
    public Appointment (string appointmentName, object location, DateTime startTime, DateTime endTime)
    {
        AppointmentName = appointmentName;

        Location = location;

        StartTime = startTime;

        EndTime= endTime;
    }
}
}