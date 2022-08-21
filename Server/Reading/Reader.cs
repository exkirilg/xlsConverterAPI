using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Server.Reading;

public class Reader : IReader
{
    private string _filePath = string.Empty;

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
    public string? ColumnsSpecification { get; private set; }
    public string? RowsSpecification { get; private set; }

    public Reader(string filePath, int headerRowOffset = default, string? columnsSpecification = default, string? rowsSpecification = default)
    {
        FilePath = filePath;
        HeaderRowOffset = headerRowOffset;
        ColumnsSpecification = columnsSpecification;
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

        foreach (var i in ParseColumnsSpecification(header))
            result.Add(i, new DataColumn(header.GetCell(i - 1).ToString()));

        return result;
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

    private int[] ParseColumnsSpecification(IRow header)
    {
        List<int> result = new();

        if (string.IsNullOrWhiteSpace(ColumnsSpecification))
            return Enumerable.Range(1, header.LastCellNum).ToArray();

        string[] cols = new string[header.LastCellNum];
        for (int i = 0; i < cols.Length; i++)
            cols[i] = header.GetCell(i).ToString()!.ToLower().Trim();

        foreach (var range in ColumnsSpecification.Split('|', StringSplitOptions.RemoveEmptyEntries))
        {
            int index = Array.IndexOf(cols, range.ToLower().Trim());
            if (index != -1)
            {
                result.Add(index + 1);
                continue;
            }

            var (start, end) = ParseRange(range, 1, header.LastCellNum);

            if (start < 1)
                throw new ArgumentException($"Starting column cannot be less then 1: {range}");

            if (end < start)
                throw new ArgumentException($"Ending column cannot be less then starting column: {range}");

            for (var i = start; i <= end; i++)
                result.Add(i);
        }

        return result.Distinct().ToArray();
    }
    private int[] ParseRowsSpecification(int firstRow, int lastRow)
    {
        List<int> result = new();

        if (string.IsNullOrWhiteSpace(RowsSpecification))
            return Enumerable.Range(firstRow, lastRow - firstRow + 1).ToArray();

        foreach (var range in RowsSpecification.Split('|', StringSplitOptions.RemoveEmptyEntries))
        {
            var (start, end) = ParseRange(range, firstRow, lastRow);

            if (start < 1)
                throw new ArgumentException($"Starting row cannot be less then 1: {range}");

            if (end < start)
                throw new ArgumentException($"Ending row cannot be less then starting row: {range}");

            for (var i = start; i <= end; i++)
                result.Add(i);
        }

        return result.Distinct().ToArray();
    }
    private (int start, int end) ParseRange(string range, int defStart, int defEnd)
    {
        (int start, int end) result = (0, 0);

        var separatorPos = range.IndexOf(':');

        if (separatorPos == -1)
            result.start = result.end = int.Parse(range);
        else
        {
            var strStart = range.Substring(0, separatorPos);
            var strEnd = range.Substring(separatorPos + 1);

            result.start = string.IsNullOrWhiteSpace(strStart) ? defStart : int.Parse(strStart);
            result.end = string.IsNullOrWhiteSpace(strEnd) ? defEnd : int.Parse(strEnd);
        }

        return result;
    }
}
