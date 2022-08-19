using Microsoft.AspNetCore.Mvc;
using System.Net.Mime;

namespace Server.Controllers;

[Route("api/[controller]")]
[ApiController]
public class ConverterController : ControllerBase
{
    private readonly string[] _availableExtensions = new string[] { ".xls", ".xlsx" };

    private readonly string _tempDirectory;

    public ConverterController(IConfiguration config)
    {
        _tempDirectory = Path.Combine(AppContext.BaseDirectory, config["TempDirectory"]);
    }

    /// <summary>
    /// Returns json file with data from uploaded xls/xlsx file
    /// </summary>
    /// <param name="file"></param>
    /// <param name="columns"></param>
    /// <param name="headerRow">Index of header row in xls/xlsx file starting with 1</param>
    /// <param name="rows">Rows indexes specification in format "1|3:10|15", starting with 1</param>
    /// <returns>Json file</returns>
    /// <response code="200">Returns json file</response>
    /// <response code="400">If uploaded file cannot be read or parameters validation error</response>
    [HttpPost("[action]")]
    [ProducesResponseType(typeof(File), StatusCodes.Status200OK, MediaTypeNames.Application.Json)]
    [ProducesResponseType(typeof(string), StatusCodes.Status400BadRequest, MediaTypeNames.Text.Plain)]
    public async Task<IActionResult> ToJson(IFormFile file, [FromQuery(Name = "column")] string[] columns, [FromQuery] int headerRow = 1, [FromQuery] string? rows = null)
    {
        if (ValidateFile(file) == false)
            return BadRequest("Uploaded file has invalid extension");

        if (headerRow < 1)
            return BadRequest("Header row value cannot be less then 1");

        string xlsFilePath = GetTempFilePath(file);

        byte[]? result = null;
        string exMessage = string.Empty;

        try
        {
            IReader reader = new Reader(xlsFilePath, headerRow - 1, columns, rows);

            await UploadToTempFile(file, xlsFilePath);
            var data = await reader.ReadFileAsync();

            var wrProcess = new WritingProcessor(WritingProcessor.JsonExtension);
            result = await wrProcess.WriteToByteArrayAsync(data);
        }
        catch (Exception ex)
        {
            exMessage = ex.Message;
            Console.WriteLine(ex);
        }
        finally
        {
            _ = Task.Run(() => DeleteTempFile(xlsFilePath));
        }

        if (result is null)
            return BadRequest($"Cannot read file: {exMessage}");

        return File(result, "text/json", GetNewFileName(file.FileName, WritingProcessor.JsonExtension));
    }

    private bool ValidateFile(IFormFile file)
    {
        if (_availableExtensions.Contains(Path.GetExtension(file.FileName)) == false)
            return false;

        return true;
    }

    private string GetTempFilePath(IFormFile file)
    {
        return Path.Combine(_tempDirectory, $"{Guid.NewGuid()}{Path.GetExtension(file.FileName)}");
    }

    private async Task UploadToTempFile(IFormFile file, string tempFilePath)
    {
        using var stream = new FileStream(tempFilePath, FileMode.Create);
        await file.CopyToAsync(stream);
    }

    private string GetNewFileName(string oldFileName, string fileExtension)
    {
        return $"{oldFileName.Substring(0, oldFileName.Length - Path.GetExtension(oldFileName).Length)}{fileExtension}";
    }

    private void DeleteTempFile(string filePath) => System.IO.File.Delete(filePath);
}
