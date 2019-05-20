using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDFGenTest.Converters
{
    public interface IConvert
    {
        void ConvertToPDF(string inputFilePath, string outputFilePath);
    }
}
