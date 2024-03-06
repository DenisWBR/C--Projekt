namespace Classes_Calendar
{
public class ExcelReaderException : Exception
{
    public ExcelReaderException(string excelFilePath, Exception innerException) : base(GenerateErrorMessage(excelFilePath, innerException)){}

    private static string GenerateErrorMessage(string excelFilePath, Exception innerException)
    {
        if (innerException is IndexOutOfRangeException)
        {
            return $"Fehler beim Einlesen der Datei: {excelFilePath}. Ein Arbeitsblatt ist außerhalb des gültigen Bereichs. Stimmt der Dateipfad?";
        }

        else if (innerException is NullReferenceException)
        {
            return $"Fehler beim Einlesen der Datei: {excelFilePath}. Das Arbeitsblatt enthält keine Daten.";
        }

        else if (innerException is IOException)
        {
            return $"Fehler beim Einlesen der Datei: {excelFilePath}. Der Zugriff auf die Datei konnte nicht durchgeführt werden.";
        }

        else if (innerException is UnauthorizedAccessException)
        {
            return $"Fehler beim Einlesen der Datei: {excelFilePath}. Es ist kein Zugriff auf die Datei möglich.";
        }

        else if (innerException is InvalidDataException)
        {
            return $"Fehler beim Einlesen der Datei: {excelFilePath}. Es wurde kein passendes Format gefunden.";
        }

        else if (innerException is FileNotFoundException)
        {
            return $"Fehler beim Einlesen der Datei: {excelFilePath}. Die angegebene Datei wurde nicht gefunden.";
        }

        else if (innerException is FormatException)
        {
            return $"Fehler beim Einlesen der Datei: {excelFilePath}. Ungültiges Format bei der Konvertierung.";
        }

        else if (innerException is ArgumentException)
        {
            return $"Fehler beim Einlesen der Datei: {excelFilePath}. Bitte geben sie eine gültige Datei ein.";
        }

        return $"Fehler beim Einlesen der Datei: {excelFilePath}. Ein unbekannter Fehler ist aufgetreten.";
    }
}
}