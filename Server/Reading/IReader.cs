namespace Server.Reading;

public interface IReader
{
    Task<DataTable> ReadFileAsync(string filePath);
    DataTable ReadFile(string fileName);
}
