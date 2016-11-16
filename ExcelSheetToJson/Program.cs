using System;
using System.Collections.Generic;
using System.IO;
using Fclp;
using Newtonsoft.Json;

namespace ExcelSheetToJson
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            var p = new FluentCommandLineParser<AppArg>();
            p.Setup(x => x.Input).As('i', "input").Required();
            p.Setup(x => x.ExcludeWorksheets).As('w', "excludeWss");
            p.Setup(x => x.ExcludeColumns).As('c', "excludeCols");
            p.Setup(x => x.Output).As('o', "output").Required();
            var result = p.Parse(args);
            if (!result.HasErrors)
            {
                var excel = new HDWorkbook();

                excel.Load(p.Object);
                File.WriteAllText(p.Object.Output, JsonConvert.SerializeObject(excel));
                Console.WriteLine("Done");
            }
            else 
                Console.WriteLine("Parameters are false");

            Console.ReadLine();
        }
    }

    internal class AppArg
    {
        public string Output { get; set; }
        public string Input { get; set; }
        public List<string> ExcludeWorksheets { get; set; }
        public List<string> ExcludeColumns { get; set; }
    }
}