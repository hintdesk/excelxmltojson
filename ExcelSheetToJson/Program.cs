using System;
using System.Collections.Generic;
using System.IO;
using Fclp;
using HDStandardLibrary.Excel;
using Newtonsoft.Json;

namespace ExcelSheetToJson
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            var p = new FluentCommandLineParser<ExcelXmlConverterArgument>();
            p.Setup(x => x.InputFilePath).As('i', "input").Required();
            p.Setup(x => x.ExcludeWorksheets).As('w', "excludeWss");
            p.Setup(x => x.ExcludeColumns).As('c', "excludeCols");
            p.Setup(x => x.OutputFilePath).As('o', "output").Required();
            p.Setup(x => x.Take).As('t', "take");

            var result = p.Parse(args);
            if (!result.HasErrors)
            {
                var excel = new HDWorkbook();

                excel.Load(p.Object);
                File.WriteAllText(p.Object.OutputFilePath, JsonConvert.SerializeObject(excel));
                Console.WriteLine("Done");
            }
            else 
                Console.WriteLine("Parameters are false");

            Console.ReadLine();
        }
    }


}