﻿using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Server.Reading;

public class Reader : IReader
{
    public async Task<DataTable> ReadFileAsync(string filePath)
    {
        return await Task.Run(() => ReadFile(filePath));
    }
    public DataTable ReadFile(string filePath)
    {
        var result = new DataTable();
        var rowList = new List<string>();

        using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read);

        var workbook = new XSSFWorkbook(stream);
        var sheet = workbook.GetSheetAt(0);
        var headerRow = sheet.GetRow(sheet.FirstRowNum);

        var cellCount = headerRow.LastCellNum;

        for (int i = 0; i < cellCount; i++)
        {
            var cell = headerRow.GetCell(i);
            if (cell is null || string.IsNullOrEmpty(cell.ToString()))
                continue;

            result.Columns.Add(cell.ToString());
        }

        for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
        {
            var row = sheet.GetRow(i);

            if (row is null || row.Cells.All(c => c.CellType == CellType.Blank))
                continue;

            for (int j = row.FirstCellNum; j < cellCount; j++)
            {
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
}
