namespace Classes_Calendar
{
public class Course
{
    //Eigenschaften
    public string CourseName {get; set; }

    public object CourseRoom {get; set; }

    public int CourseColumnIndex {get; set; }

    //Konstruktor
    public Course(string courseName, object courseRoom, int courseColumnIndex)
    {
        CourseName = courseName;

        CourseRoom = courseRoom;

        CourseColumnIndex = courseColumnIndex;
    }
}
}