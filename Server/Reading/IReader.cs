namespace Server.Reading;

public interface IReader
{
    Task<DataTable> ReadFileAsync(string filePath, int headerRowOffset, string[] specifiedCols);
    DataTable ReadFile(string fileName, int headerRowOffset, string[] specifiedCols);
}
