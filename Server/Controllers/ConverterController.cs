using Microsoft.AspNetCore.Mvc;
using System.Net.Mime;

namespace Server.Controllers;

[Route("api/[controller]")]
[ApiController]
public class ConverterController : ControllerBase
{
    private readonly string[] _availableExtensions = new string[] { ".xls", ".xlsx" };

    private readonly string _tempDirectory;
    private readonly IReader _reader;

    public ConverterController(IConfiguration config, IReader reader)
    {
        _reader = reader;
        _tempDirectory = Path.Combine(AppContext.BaseDirectory, config["TempDirectory"]);
    }

    /// <summary>
    /// Returns json file with data from uploaded xls/xlsx file
    /// </summary>
    /// <param name="file"></param>
    /// <param name="columns"></param>
    /// <returns>Json file</returns>
    /// <response code="200">Returns json file</response>
    /// <response code="400">If uploaded file is invalid</response>
    [HttpPost("[action]")]
    [ProducesResponseType(typeof(File), StatusCodes.Status200OK, MediaTypeNames.Application.Json)]
    [ProducesResponseType(typeof(string), StatusCodes.Status400BadRequest, MediaTypeNames.Text.Plain)]
    public async Task<IActionResult> ToJson(IFormFile file, [FromQuery(Name = "column")] string[]? columns = null)
    {
        if (ValidateFile(file) == false)
            return BadRequest("Uploaded file has invalid extension");

        var xlsFilePath = await UploadToTempFile(file);
        var data = await _reader.ReadFileAsync(xlsFilePath, NormalizeSpecifiedColumns(columns));
        
        var wrProcess = new WritingProcessor(WritingProcessor.JsonExtension);
        var byteArray = await wrProcess.WriteToByteArrayAsync(data);

        _ = Task.Run(() => DeleteTempFile(xlsFilePath));

        return File(byteArray, "text/json", GetNewFileName(file.FileName, WritingProcessor.JsonExtension));
    }

    private bool ValidateFile(IFormFile file)
    {
        if (_availableExtensions.Contains(Path.GetExtension(file.FileName)) == false)
            return false;

        return true;
    }

    private async Task<string> UploadToTempFile(IFormFile file)
    {
        var filePath = Path.Combine(_tempDirectory, $"{Guid.NewGuid()}{Path.GetExtension(file.FileName)}");
        
        using var stream = new FileStream(filePath, FileMode.Create);
        await file.CopyToAsync(stream);

        return filePath;
    }

    private string GetNewFileName(string oldFileName, string fileExtension)
    {
        return $"{oldFileName.Substring(0, oldFileName.Length - Path.GetExtension(oldFileName).Length)}{fileExtension}";
    }

    private void DeleteTempFile(string filePath) => System.IO.File.Delete(filePath);

    private string[]? NormalizeSpecifiedColumns(string[]? specifiedCols)
    {
        if (specifiedCols is null)
            return null;
        
        return specifiedCols.Select(c => c.ToLower().Trim()).ToArray();
    }
}
