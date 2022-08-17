namespace Server.Writing;

public class WritingProcessor
{
    public const string JsonExtension = ".json";

    private readonly IWriter _writer;

    public WritingProcessor(string fileExtension)
    {
        _writer = fileExtension switch
        {
            JsonExtension => new WriterJson(),
            _ => throw new ArgumentException($"Invalid file extension: {fileExtension}."),
        };
    }

    public async Task<byte[]> WriteToByteArrayAsync(DataTable data) => await _writer.WriteToByteArrayAsync(data);
}
