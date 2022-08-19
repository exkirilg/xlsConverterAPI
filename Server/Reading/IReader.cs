namespace Server.Reading;

public interface IReader
{
    Task<DataTable> ReadFileAsync();
}
