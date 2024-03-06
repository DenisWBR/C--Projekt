namespace Classes_Calendar
{

public class TimeSlot
{
    //Eigenschaften
    public DateTime StartTime {get; set; }

    public DateTime EndTime {get; set; }

    public int TimeSlotRowIndex {get; set; }

    //Konstruktor
    public TimeSlot (DateTime startTime, DateTime endTime, int timeSlotRowIndex)
    {
        StartTime = startTime;
        
        EndTime = endTime;

        TimeSlotRowIndex = timeSlotRowIndex;
    }
}
}