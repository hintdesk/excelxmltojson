using System;
using System.Collections.Generic;
using System.Dynamic;
using HDStandardLibrary.Excel;

namespace ExcelSheetToJson
{

    internal class HDWorkbook
    {
        public HDWorkbook()
        {
            Worksheets = new List<ExcelXmlConverterWorksheet>();
        }

        public IList<ExcelXmlConverterWorksheet> Worksheets { get; set; }

        public void Load(ExcelXmlConverterArgument appArg)
        {
            if (appArg.InputFilePath.EndsWith(".xml"))
            {
                var xmlLoader = new ExcelXmlConverter();
                Worksheets = xmlLoader.Load(appArg);
            }
            else 
                throw new NotSupportedException("Only .xml file is supported");
        }
    }
}