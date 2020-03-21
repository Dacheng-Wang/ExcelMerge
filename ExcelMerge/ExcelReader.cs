using System.Collections.Generic;
using NPOI.SS.UserModel;

namespace ExcelMerge
{
    internal class ExcelReader
    {
        internal static IEnumerable<ExcelRow> Read(ISheet sheet)
        {
            //bool isRowEmpty = false;
            var actualRowIndex = 0;
            //var lastRow = sheet.LastRowNum;
            //for (int rowIndex = sheet.LastRowNum; rowIndex>0; rowIndex--)
            //{
            //    var row = sheet.GetRow(rowIndex);
            //    if (row == null) isRowEmpty = true;
            //    else
            //    {
            //        foreach (var cell in row.Cells)
            //        {
            //            if (cell.CellType != CellType.Blank) isRowEmpty = true;
            //        }
            //    }
            //    if (isRowEmpty) lastRow--;
            //    else break;
            //}
            for (int rowIndex = 0; rowIndex <= sheet.PhysicalNumberOfRows; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                if (row == null)
                    continue;

                var cells = new List<ExcelCell>();
                for (int columnIndex = 0; columnIndex < row.PhysicalNumberOfCells; columnIndex++)
                {
                    var cell = row.GetCell(columnIndex);
                    var stringValue = ExcelUtility.GetCellStringValue(cell);

                    if (cell != null && cell.CellType != CellType.Blank) cells.Add(new ExcelCell(stringValue, columnIndex, rowIndex));
                }

                yield return new ExcelRow(actualRowIndex++, cells);
            }
        }
    }
}
