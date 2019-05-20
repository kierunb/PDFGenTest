using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDFGenTest.Converters
{
    public class AsposeWord : IConvert
    {
        public void ToConvertPDF(string inputFilePath, string outputFilePath)
        {
            Document doc = new Document(inputFilePath);
            doc.Save(outputFilePath);
        }
    }
}
