namespace Server.Writing;

public interface IWriter
{
    Task<byte[]> WriteToByteArrayAsync(DataTable data);
}
