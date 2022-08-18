namespace Server.Reading;

public interface IReader
{
    Task<DataTable> ReadFileAsync(string filePath, string[]? specifiedCols);
    DataTable ReadFile(string fileName, string[]? specifiedCols);
}
