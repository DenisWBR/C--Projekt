using DocumentFormat.OpenXml.Spreadsheet;
using OfficeOpenXml;

namespace Classes_Calendar
{

/// <summary>
/// Klasse, die für das Auslesen des Excel-Dokumentes zuständig ist.
/// </summary>
public class ExcelReader
{
    //Eigenschaften
    string ExcelFilePath {get; }

    //Konstruktor
    public ExcelReader (string excelFilePath)
    {
        ExcelFilePath = excelFilePath;
    }

    /// <summary>
    /// Liest aus, in welcher Zeile der Tabelle die Kursbezeichnungen stehen.
    /// </summary>
    /// <returns></returns>
    public int ReadCoursesRow()
    {
        try
        {    
            using(var package = new ExcelPackage(new System.IO.FileInfo(ExcelFilePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;
                int columnIndex = 1;

                for (int subjectRowIndex = 1; subjectRowIndex <= rowCount; subjectRowIndex++)
                {
                    string? courseString = worksheet.Cells[subjectRowIndex, columnIndex].Value as String;

                    if(courseString == "Kurs")
                    {
                        int courseRow = subjectRowIndex;
                        return courseRow;
                    }
                }
            }
        return 0;
        }

        catch(Exception ex) 
        {
            throw new ExcelReaderException(ExcelFilePath, ex);
        }   
    }

    /// <summary>
    /// Liest aus der vorliegenden Excel-Datei die vorhandenen Studiengänge aus und speichert den zugehörigen Spaltenindex ab.
    /// </summary>
    public List<Course> ReadCourses()
    {
        List<Course> courses = new List<Course>();
        
        try
        {    
            using(var package = new ExcelPackage(new System.IO.FileInfo(ExcelFilePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int courseLength;
                int rowCount = worksheet.Dimension.Rows;

                for (int courseColumnIndex = 1; courseColumnIndex <= 25; courseColumnIndex++)
                {
                    string? courseName = worksheet.Cells[4, courseColumnIndex].Value as String;
                    string? courseGroup = worksheet.Cells[5, courseColumnIndex].Value as String;
                    object courseRoom = worksheet.Cells[4, courseColumnIndex+1].Value;

                    if (courseName != null)
                    {
                        string course = courseName.ToString();
                        string group = courseGroup.ToString();
                        courseName = $"{course} {group}";
                        courseLength = course.Length;   
                    }

                    else
                    {
                        string course = "NULL";
                        courseLength = course.Length;
                    }

                    if (courseLength > 4)
                    {   
                        Course course = new Course(courseName, courseRoom, courseColumnIndex);
                        courses.Add(course);
                    }
                }
            }

            return courses;
        }

        catch(Exception ex)
        {
            throw new ExcelReaderException(ExcelFilePath, ex);
        }
    }


    /// <summary>
    /// Liest die Vorlesungen des Vorlesungsplans anhand des ausgewählten Spaltenindexes aus und Speichert die zugehörigen Zeilenindexe. 
    /// </summary>
    public List<Subject> ReadSubjects(int courseColumn, int courseRow)
    {
        List<Subject> subjects = new List<Subject>();

        try
        {
            using (var package = new ExcelPackage(new System.IO.FileInfo(ExcelFilePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;
                object location = worksheet.Cells[courseRow, courseColumn+1].Value;

                for (int subjectRowIndex = 6; subjectRowIndex <= rowCount; subjectRowIndex++)
                {          
                    string? subjectName = worksheet.Cells[subjectRowIndex, courseColumn].Value as String;

                    if (subjectName != null)
                    {
                        Subject subject = new Subject(subjectName, location, subjectRowIndex);
                        subjects.Add(subject);  
                    }                 
                }
            
                return subjects;
            }
        }
        catch (Exception ex)
        {
            throw new ExcelReaderException(ExcelFilePath, ex);
        }
    }


    /// <summary>
    /// Liest die vorhandenen Zeitstempel aus und speichert die zugehörigen Zeilen ab.
    /// </summary>
    public List<TimeSlot> ReadTimes()
    {
        List<TimeSlot> timeSlots = new List<TimeSlot>();

        try
        {
            using (var package = new ExcelPackage(new System.IO.FileInfo(ExcelFilePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int timeSlotRowIndex = 6; timeSlotRowIndex <= rowCount; timeSlotRowIndex++ )
                {
                    var dateValue = worksheet.Cells[timeSlotRowIndex,1].GetValue<DateTime>();
                    var timeValueStart = worksheet.Cells[timeSlotRowIndex,3].GetValue<DateTime>();
                    var timeValueEnd = worksheet.Cells[timeSlotRowIndex,4].GetValue<DateTime>();

                    DateTime startTime = new DateTime (dateValue.Year, dateValue.Month, dateValue.Day, timeValueStart.Hour, timeValueStart.Minute, timeValueStart.Second);
                    DateTime endTime = new DateTime (dateValue.Year, dateValue.Month, dateValue.Day, timeValueEnd.Hour, timeValueEnd.Minute, timeValueEnd.Second);

                    TimeSlot timeSlot = new TimeSlot(startTime, endTime, timeSlotRowIndex);
                    timeSlots.Add(timeSlot);      
                }
            }

            return timeSlots;
        }

        catch(Exception ex)
        {
            throw new ExcelReaderException(ExcelFilePath, ex);
        }
    } 
}
}