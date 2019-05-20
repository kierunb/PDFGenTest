using Aspose.Words;
using PDFGenTest.Converters;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDFGenTest
{
    class Program
    {
        static void Main(string[] args)
        {
            string inputDirectory = Path.Combine(Environment.CurrentDirectory, "input");
            string outputDirectory = Path.Combine(Environment.CurrentDirectory, "output");

            var stopwatch = new Stopwatch();

            var files = Directory.GetFiles(inputDirectory);

            if (!Directory.Exists(outputDirectory))
                Directory.CreateDirectory(outputDirectory);

            var converters = new Dictionary<string, IConvert>()
            {
                { nameof(AsposeWord), new AsposeWord() },
            }; 
            
            Console.WriteLine($"Processing of {files.Length} files using {converters.Count} converters started.\n");
            int counter = 0;
            

            foreach (var item in converters)
            {
                string converterName = item.Key;
                Console.WriteLine($"Converter used: {converterName}");
                IConvert converter = item.Value;

                string outputFolderForConverter = Path.Combine(outputDirectory, converterName);
                if (!Directory.Exists(outputFolderForConverter))
                    Directory.CreateDirectory(outputFolderForConverter);

                stopwatch.Start();
                foreach (var file in files)
                {
                    Console.WriteLine($"Processing file {++counter}/{files.Length}: {Path.GetFileName(file)}...");
                    string outputFileName = $"{Path.GetFileNameWithoutExtension(file)}.pdf";
                    string outputFilePath = Path.Combine(outputFolderForConverter, outputFileName);

                    try
                    {
                        converter.ConvertToPDF(file, outputFilePath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Exception: {ex.Message}");
                    }

                }
                stopwatch.Stop();
                Console.WriteLine($"Processing of {files.Length} files with {converterName} converter completed in: {stopwatch.Elapsed.Seconds} seconds.\n");

            }

            Console.ReadKey();
        }




    }
}
