using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Server.Reading;

public class Reader : IReader
{
    private string _filePath = string.Empty;
    private string[]? _specifiedColumns = null;

    public string FilePath
    {
        get => _filePath;
        private set
        {
            if (File.Exists(value) == false)
            {
                using (File.Create(value)) { };
                File.Delete(value);
            }

            _filePath = value;
        }
    }
    public int HeaderRowOffset { get; private set; }
    public string[]? SpecifiedColumns
    {
        get => _specifiedColumns;
        private set
        {
            if (value is null)
            {
                _specifiedColumns = null;
                return;
            }

            _specifiedColumns = value.Select(c => c.ToLower().Trim()).ToArray();
        }
    }
    public string? RowsSpecification { get; private set; }

    public Reader(string filePath, int headerRowOffset = 0, string[]? specifiedColumns = null, string? rowsSpecification = null)
    {
        FilePath = filePath;
        HeaderRowOffset = headerRowOffset;
        SpecifiedColumns = specifiedColumns;
        RowsSpecification = rowsSpecification;
    }

    public async Task<DataTable> ReadFileAsync() => await Task.Run(() => ReadFile());
    
    private DataTable ReadFile()
    {
        var result = new DataTable();
        var rowList = new List<string>();

        using var stream = new FileStream(FilePath, FileMode.Open, FileAccess.Read);

        var workbook = new XSSFWorkbook(stream);
        var sheet = workbook.GetSheetAt(0);

        var headerRowIndex = sheet.FirstRowNum + HeaderRowOffset;
        var headerRow = sheet.GetRow(headerRowIndex);

        var cellCount = headerRow.LastCellNum;
        
        for (int i = 0; i < cellCount; i++)
        {
            var cell = headerRow.GetCell(i);
            
            if (cell is null || string.IsNullOrEmpty(cell.ToString()))
                continue;

            if (UploadColumn(cell.ToString()!) == false)
                continue;

            result.Columns.Add(cell.ToString());
        }

        int firstRowIndex = headerRowIndex + 1;
        int[] rows = ParseRowsSpecification(firstRowIndex + 1, sheet.PhysicalNumberOfRows);

        for (int i = firstRowIndex; i <= sheet.PhysicalNumberOfRows; i++)
        {
            if (rows.Contains(i + 1) == false)
                continue;

            var row = sheet.GetRow(i);

            if (row is null || row.Cells.All(c => c.CellType == CellType.Blank))
                continue;

            for (int j = row.FirstCellNum; j < cellCount; j++)
            {
                if (UploadColumn(headerRow.GetCell(j).ToString()!) == false)
                    continue;

                var cell = row.GetCell(j);

                if (cell is null | string.IsNullOrWhiteSpace(cell!.ToString()))
                    continue;

                rowList.Add(cell.ToString()!);
            }

            if (rowList.Any())
            {
                result.Rows.Add(rowList.ToArray());
                rowList.Clear();
            }
        }

        return result;
    }

    private bool UploadColumn(string columnName)
    {
        if (SpecifiedColumns is null || SpecifiedColumns.Any() == false)
            return true;

        return SpecifiedColumns.Contains(columnName.ToLower());
    }

    private int[] ParseRowsSpecification(int firstRow, int lastRow)
    {
        List<int> result = new();

        if (string.IsNullOrWhiteSpace(RowsSpecification))
            return Enumerable.Range(firstRow, lastRow - firstRow + 1).ToArray();

        foreach (var range in RowsSpecification.Split('|', StringSplitOptions.RemoveEmptyEntries))
        {
            int start, end;
            var separatorPos = range.IndexOf(':');

            if (separatorPos == -1)
                start = end = int.Parse(range);
            else
            {
                var strStart = range.Substring(0, separatorPos);
                var strEnd = range.Substring(separatorPos + 1);

                start = string.IsNullOrWhiteSpace(strStart) ? firstRow : int.Parse(strStart);
                end = string.IsNullOrWhiteSpace(strEnd) ? lastRow : int.Parse(strEnd);
            }

            if (start < 1)
                throw new ArgumentException("Starting row cannot be less then 1");

            if (end < start)
                throw new ArgumentException("Ending row cannot be less then starting row");

            for (var i = start; i <= end; i++)
                result.Add(i);
        }

        return result.Distinct().ToArray();
    }
}
