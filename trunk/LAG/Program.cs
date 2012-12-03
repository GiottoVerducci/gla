using System;
using System.IO;
using System.Linq;

namespace GLA
{
    public struct Arguments
    {
        public string InputFileName;
        public string OutputFileName;
    }

    class Program
    {
        static void Main(string[] args)
        {
            Arguments arguments;

            if (!TryGetArguments(args, out arguments))
            {
                WriteUsage();
                return;
            }

            var filename = //@"c:\Users\vripoll\Documents\Visual Studio 2012\Projects\LAG\Book Ophidien.xlsx";
                //@"..\..\..\Book Ophidien.xlsx";
                arguments.InputFileName;
            
            //var army = ExcelLoader.ReadArmy(filename);
            var armies = ExcelLoader.LoadAll(@"..\..\..\");

            foreach (var line in Warnings.GetSummary())
                Console.WriteLine(line);

            var modificationDate = File.GetLastWriteTime(filename);

            //var footer = FormatProperty(GetGeneralProperty(army, "Pied de page"), modificationDate);
            //var headerLeft = FormatProperty(GetGeneralProperty(army, "Titre"), modificationDate);
            //var headerRight = FormatProperty(GetGeneralProperty(army, "Haut de page droit"), modificationDate);
            //var watermarkPath = GetGeneralProperty(army, "Watermark");
            //int watermarkOpacity;
            //if (!Int32.TryParse(GetGeneralProperty(army, "Watermark opacity"), out watermarkOpacity))
            //    watermarkOpacity = 50;

            //Pdf.Write(army, arguments.OutputFileName, footer, headerLeft, headerRight, watermarkPath, watermarkOpacity);

            var footer = FormatProperty("Modifié le {Date}", modificationDate);
            var headerRight = FormatProperty("Confédération du Dragon Rouge Française", modificationDate);
            var watermarkPath = @"..\..\..\dragon.jpeg";
            int watermarkOpacity;
            if (!Int32.TryParse("30", out watermarkOpacity))
                watermarkOpacity = 50;

            foreach (var army in armies)//.Where(a => a.Key._name.Contains("Immortel")))
            {
                var headerLeft = FormatProperty("Livre d'armée " + army.Key._name, modificationDate);
                Pdf.Write(army.Value, army.Key, arguments.OutputFileName + army.Key._name + ".pdf", footer, headerLeft, headerRight, watermarkPath, watermarkOpacity);
                break;
            }
        }

        private static string FormatProperty(string generalProperty, DateTime modificationDate)
        {
            var result = generalProperty.Replace("{Date}", "{0}");
            return string.Format(result, modificationDate.ToShortDateString());
        }

        private static string GetGeneralProperty(Army army, string propertyName)
        {
            string result;
            if (army.GeneralProperties.TryGetValue(propertyName, out result))
                return result;
            Console.WriteLine("Propriété générale '{0}' non trouvée.", propertyName);
            return string.Empty;
        }

        private static bool TryGetArguments(string[] args, out Arguments arguments)
        {
            arguments = new Arguments();
            if (args.Length < 2 || args.Length > 11)
                return false;

            // default values
            arguments.InputFileName = args[0];
            arguments.OutputFileName = args[1];

            return args.Length == 2;
        }

        private static void WriteUsage()
        {
            Console.WriteLine("Invalid arguments.");
            Console.WriteLine("Expected arguments: <filename>.xlsx <outputname>.pdf");
        }
    }
}
