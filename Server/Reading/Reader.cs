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
        DataTable result = new();
        List<string> rowList = new();

        using FileStream stream = new(FilePath, FileMode.Open, FileAccess.Read);

        XSSFWorkbook workbook = new(stream);
        ISheet sheet = workbook.GetSheetAt(0);

        var columns = ReadHeader(sheet.GetRow(sheet.FirstRowNum + HeaderRowOffset));
        var rows = ReadDataRows(sheet, columns.Keys.Select(c => c - 1).ToArray());

        result.Columns.AddRange(columns.Values.ToArray());
        foreach (var row in rows)
        {
            result.Rows.Add(row);
        }

        return result;
    }

    private Dictionary<int, DataColumn> ReadHeader(IRow header)
    {
        Dictionary<int, DataColumn> result = new();

        for (int i = 0; i < header.LastCellNum; i++)
        {
            string colName = header.GetCell(i).ToString()!;
            
            if (ColumnIsSpecified(i + 1, colName))
            {
                result.Add(i + 1, new DataColumn(colName));
            }           
        }

        return result;
    }

    private bool ColumnIsSpecified(int columnOneBasedIndex, string columnName)
    {
        if (SpecifiedColumns is null || SpecifiedColumns.Any() == false)
            return true;

        return SpecifiedColumns
            .Where(
                c => c.Equals(columnName?.ToLower()) ||
                (int.TryParse(c, out int specInt) && specInt == columnOneBasedIndex))
            .Any();
    }

    private List<string[]> ReadDataRows(ISheet sheet, int[] columnsZeroBasedIndexes)
    {
        List<string[]> result = new();

        int firstDataRowIndex = sheet.FirstRowNum + HeaderRowOffset + 1;
        int[] rows = ParseRowsSpecification(firstDataRowIndex + 1, sheet.PhysicalNumberOfRows);

        List<string> rowData = new();

        for (int i = firstDataRowIndex; i <= sheet.PhysicalNumberOfRows; i++)
        {
            if (rows.Contains(i + 1) == false)
                continue;

            var row = sheet.GetRow(i);

            if (row is null || row.Cells.All(c => c.CellType == CellType.Blank))
                continue;

            for (int j = row.FirstCellNum; j < row.LastCellNum; j++)
            {
                if (columnsZeroBasedIndexes.Contains(j) == false)
                    continue;

                rowData.Add(row.GetCell(j).ToString()!);
            }

            result.Add(rowData.ToArray());
            rowData.Clear();
        }

        return result;
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
