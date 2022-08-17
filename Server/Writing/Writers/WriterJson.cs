using Newtonsoft.Json;
using System.Text;

namespace Server.Writing.Writers;

public class WriterJson : IWriter
{
    public async Task<byte[]> WriteToByteArrayAsync(DataTable data)
    {
        return await Task.Run(() => Encoding.UTF8.GetBytes(SerializeDataTableToJson(data)));
    }

    private string SerializeDataTableToJson(DataTable data) => JsonConvert.SerializeObject(data);
}
