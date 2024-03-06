namespace Classes_Calendar
{
    public class Subject
    {
        //Eigenschaften
        public string SubjectName { get; set; }

        public object Location { get; set; }

        public int SubjectRowIndex { get; set; }

        public int SubjectRowIndexStart { get; set; }

        public int SubjectRowIndexEnd { get; set; }

        //Konstruktor
        public Subject(string subjectName, object location, int subjectRowIndex)
        {
            SubjectName = subjectName;

            Location = location;

            SubjectRowIndex = subjectRowIndex;

            SubjectRowIndexStart = subjectRowIndex;

            SubjectRowIndexEnd = subjectRowIndex;
        }
    }
}

