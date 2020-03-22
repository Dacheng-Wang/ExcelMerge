using System.Collections.Generic;
using System.IO;
using NPOI.SS.UserModel;

namespace ExcelMerge
{
    public class ExcelWorkbook
    {
        public Dictionary<string, ExcelSheet> Sheets { get; private set; }
        public string WorkbookPath { get; private set; }
        public ExcelWorkbook()
        {
            Sheets = new Dictionary<string, ExcelSheet>();
        }

        public static ExcelWorkbook Create(string path, ExcelSheetReadConfig config)
        {
            if (Path.GetExtension(path) == ".csv")
                return CreateFromCsv(path, config);

            if (Path.GetExtension(path) == ".tsv")
                return CreateFromTsv(path, config);

            var srcWb = WorkbookFactory.Create(path);
            var wb = new ExcelWorkbook();
            for (int i = 0; i < srcWb.NumberOfSheets; i++)
            {
                var srcSheet = srcWb.GetSheetAt(i);
                wb.Sheets.Add(srcSheet.SheetName, ExcelSheet.Create(srcSheet, config));
            }
            wb.WorkbookPath = path;
            return wb;
        }

        public static IEnumerable<string> GetSheetNames(ExcelWorkbook wb)
        {
            if (Path.GetExtension(wb.WorkbookPath) == ".csv")
            {
                yield return "csv";
            }
            else if (Path.GetExtension(wb.WorkbookPath) == ".tsv")
            {
                yield return "tsv";
            }
            else
            {
                foreach (KeyValuePair<string, ExcelSheet> pair in wb.Sheets)
                {
                    yield return pair.Key;
                }
            }
        }

        private static ExcelWorkbook CreateFromCsv(string path, ExcelSheetReadConfig config)
        {
            var wb = new ExcelWorkbook();
            wb.Sheets.Add("csv", ExcelSheet.CreateFromCsv(path, config));

            return wb;
        }

        private static ExcelWorkbook CreateFromTsv(string path, ExcelSheetReadConfig config)
        {
            var wb = new ExcelWorkbook();
            wb.Sheets.Add("tsv", ExcelSheet.CreateFromTsv(path, config));

            return wb;
        }
    }
}
