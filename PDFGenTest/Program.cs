using Aspose.Words;
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


            Console.WriteLine($"Processing of {files.Length} files started.");
            int counter = 0;
            stopwatch.Start();

            foreach (var file in files)
            {
                Console.WriteLine($"Processing file {++counter}/{files.Length}: {Path.GetFileName(file)}...");
                string outputFileName = $"{Path.GetFileNameWithoutExtension(file)}.pdf";
                string outputFilePath = Path.Combine(outputDirectory, outputFileName);

                AsposeWordConvert(file, outputFilePath);
            }

            stopwatch.Stop();
            Console.WriteLine($"Processing of {files.Length} completed in: {stopwatch.Elapsed.Seconds} seconds.");
            Console.ReadKey();
        }


        static void AsposeWordConvert(string inputFilePath, string outputFilePath)
        {
            try
            {
                Document doc = new Document(inputFilePath);
                doc.Save(outputFilePath);
            }
            catch (Exception ex)
            {

                Console.WriteLine($"Exception: {ex.Message}");
            }


        }



    }
}
