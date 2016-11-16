using System.Collections.Generic;
using System.Dynamic;

namespace ExcelSheetToJson
{
    internal class HDWorksheet
    {
        public HDWorksheet()
        {
            Rows = new List<dynamic>();
        }

        public string Name { get; set; }
        public dynamic Rows { get; set; }
    }

    internal class HDWorkbook
    {
        public HDWorkbook()
        {
            Worksheets = new List<HDWorksheet>();
        }

        public IList<HDWorksheet> Worksheets { get; set; }

        public void Load(AppArg appArg)
        {
            if (appArg.Input.EndsWith(".xml"))
            {
                var xmlLoader = new XmlLoader();
                Worksheets = xmlLoader.Load(appArg);
            }
        }
    }
}