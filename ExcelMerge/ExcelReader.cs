using System.Collections.Generic;
using NPOI.SS.UserModel;
using System;

namespace ExcelMerge
{
    internal class ExcelReader
    {
        internal static IEnumerable<ExcelRow> Read(ISheet sheet)
        {
            var actualRowIndex = 0;
            int maxColumn = GetMaxColumn(sheet);
            int maxRow = GetMaxRow(sheet);
            for (int rowIndex = 0; rowIndex <= maxRow; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                var cells = new List<ExcelCell>();
                if (row == null || row.Cells.Count == 0)
                {
                    yield return new ExcelRow(actualRowIndex++, cells);
                    continue;
                }
                int firstColumn = row.FirstCellNum;
                int currentRowMaxColumn = Math.Min(row.LastCellNum - 1, maxColumn);
                for (int columnIndex = 0; columnIndex <= currentRowMaxColumn; columnIndex++)
                {
                    var cell = row.GetCell(columnIndex);
                    var stringValue = ExcelUtility.GetCellStringValue(cell);
                    cells.Add(new ExcelCell(stringValue, columnIndex, rowIndex));
                }

                yield return new ExcelRow(actualRowIndex++, cells);
            }
        }
        internal static int GetMaxColumn(ISheet sheet)
        {
            int maxColumn = 0;
            //Mark empty columns so it won't be added to cells to get rendered/compared later
            for (int rowIndex = sheet.FirstRowNum; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                if (row == null || row.Cells.Count == 0)
                    continue;
                int firstColumn = row.FirstCellNum;
                for (int columnIndex = row.LastCellNum - 1; columnIndex >= firstColumn; columnIndex--)
                {
                    if (row.GetCell(columnIndex) != null && row.GetCell(columnIndex).CellType != CellType.Blank)
                    {
                        maxColumn = Math.Max(maxColumn, columnIndex);
                        break;
                    }
                }
            }
            return maxColumn;
        }
        internal static int GetMaxRow(ISheet sheet)
        {
            int maxRow = 0;
            //Mark empty rows so it won't be added to cells to get rendered/compared later
            for (int rowIndex = sheet.LastRowNum; rowIndex >= sheet.FirstRowNum; rowIndex--)
            {
                var row = sheet.GetRow(rowIndex);
                if (row == null || row.Cells.Count == 0)
                    continue;
                int firstRow = row.FirstCellNum;
                for (int columnIndex = row.FirstCellNum; columnIndex <= row.LastCellNum - 1; columnIndex++)
                {
                    if (row.GetCell(columnIndex) != null && row.GetCell(columnIndex).CellType != CellType.Blank)
                    {
                        maxRow = Math.Max(maxRow, rowIndex);
                        break;
                    }
                }
            }
            return maxRow;
        }
    }
}