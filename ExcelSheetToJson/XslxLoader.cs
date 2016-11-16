using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using ClosedXML.Excel;

namespace ExcelSheetToJson
{
    internal class XslxLoader
    {
        private IList<HDWorksheet> Load(string fullFilePath)
        {
            IList<HDWorksheet> worksheets = new List<HDWorksheet>();
            var workBook = new XLWorkbook(fullFilePath);
            foreach (var xlWorksheet in workBook.Worksheets)
            {
                worksheets.Add(LoadWorksheet(xlWorksheet));
            }
            return worksheets;
        }

        private HDWorksheet LoadWorksheet(IXLWorksheet xlWorksheet)
        {
            var worksheet = new HDWorksheet();
            var columNames = new List<string>();
            for (var i = 0; i < xlWorksheet.ColumnCount(); i++)
            {
                columNames.Add(xlWorksheet.Rows().ElementAt(0).Cells().ElementAt(i).GetString());
            }

            for (var i = 1; i < xlWorksheet.RowCount(); i++)
            {
                worksheet.Rows.Add(LoadRow(columNames, xlWorksheet.Rows().ElementAt(i)));
            }
            return worksheet;
        }

        private dynamic LoadRow(List<string> columNames, IXLRow row)
        {
            var expandoObject = new ExpandoObject() as IDictionary<string, object>;
            for (var i = 0; i < columNames.Count; i++)
            {
                expandoObject.Add(columNames[i], row.Cells().ElementAt(i).GetString());
            }
            return expandoObject;
        }
    }
}