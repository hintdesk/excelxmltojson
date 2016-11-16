using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Drawing;

namespace ExcelSheetToJson
{
    internal class XmlLoader
    {
        private XmlUtil xmlUtil;

        private IList<XElement> Rows { get; set; }
        private int ColumnCount { get; set; }
        private int RowCount { get; set; }

        private List<string> ExcludedWorksheets { get; set; }
        private List<string> ExcludeColumns { get; set; }

        public IList<HDWorksheet> Load(AppArg appArg)
        {
            ExcludedWorksheets = appArg.ExcludeWorksheets;
            ExcludeColumns = appArg.ExcludeColumns;

            IList<HDWorksheet> worksheets = new List<HDWorksheet>();
            xmlUtil = new XmlUtil(File.ReadAllText(appArg.Input));
            var xWorkSheets = xmlUtil.GetElements("Worksheet");
            foreach (var xWorkSheet in xWorkSheets)
            {
                var worksheet = LoadWorksheet(xWorkSheet);
                if (worksheet != null)
                    worksheets.Add(worksheet);
            }
            return worksheets;
        }

        private HDWorksheet LoadWorksheet(XElement xWorkSheet)
        {
            var worksheet = new HDWorksheet();
            var xName = this.xmlUtil.GetAttributeValue(xWorkSheet, "Name");
            if (xName != null)
            {
                worksheet.Name = xName;
                if (ExcludedWorksheets.Contains(xName, StringComparer.OrdinalIgnoreCase))
                    return null;
            }
            Rows = xmlUtil.GetRelativeElements(xWorkSheet, "Table.Row");
            ColumnCount = GetColumnCount();
            RowCount = GetRowCount();

            var columNames = new List<string>();
            var rowValues = GetRowValues(0);
            for (var columnIndex = 0; columnIndex < ColumnCount; columnIndex++)
            {
                columNames.Add(Normalize(rowValues[columnIndex]));
            }

            for (var rowIndex = 1; rowIndex < RowCount; rowIndex++)
            {
                worksheet.Rows.Add(LoadRow(columNames, rowIndex));
            }
            return worksheet;
        }

        private string Normalize(string rowValue)
        {
            return
                rowValue.Replace(" ", "")
                    .Replace(".", "")
                    .Replace("-", "")
                    .Replace("(", "")
                    .Replace(")", "")
                    .Replace("=", "")
                    .Replace(",", "")
                    .Replace("/", "")
                    .Replace("\n", "");
        }

        private dynamic LoadRow(List<string> columNames, int rowIndex)
        {
            var expandoObject = new ExpandoObject() as IDictionary<string, object>;
            var rowValues = GetRowValues(rowIndex);
            for (var columnIndex = 0; columnIndex < columNames.Count; columnIndex++)
            {
                if (columnIndex < rowValues.Count &&
                    !ExcludeColumns.Contains(columNames[columnIndex], StringComparer.OrdinalIgnoreCase))
                {
                    expandoObject.Add(columNames[columnIndex], rowValues[columnIndex]);
                }
            }
            return expandoObject;
        }

        private IList<string> GetRowValues(int rowIndex)
        {
            IList<string> result = new List<string>();
            var cellElements = this.xmlUtil.GetRelativeElements(Rows[rowIndex], "Cell");
            for (int index = 0; index < cellElements.Count; index++)
            {
                var indexReal = this.xmlUtil.GetAttributeValue(cellElements[index], "Index");
                bool isAdd = !string.IsNullOrWhiteSpace(indexReal);
                int indexRealInt=index;
                if (isAdd && !string.IsNullOrWhiteSpace(indexReal))
                    isAdd = int.TryParse(indexReal, out indexRealInt) && indexRealInt != index;
                
                if (isAdd)
                {
                    for (int subIndex = index; subIndex < indexRealInt-1; subIndex++)
                    {
                        result.Add(null);
                    }
                }

                result.Add(this.xmlUtil.GetRelativeElement(cellElements[index], "Data").Value);
            }
            return result;
        }

        private int GetRowCount()
        {
            return Rows.Count();
        }


        private int GetColumnCount()
        {
            return this.xmlUtil.GetRelativeElements(Rows[0], "Cell.Data").Count();
        }
    }
}