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
        public void ConvertToPDF(string inputFilePath, string outputFilePath)
        {
            Document doc = new Document(inputFilePath);
            doc.UpdateFields();
            doc.UpdatePageLayout();
            doc.UpdateTableLayout();
            doc.UpdateListLabels();
            doc.Save(outputFilePath);
        }

        public void ConvertToPDF(string[] inputFilePath, string[] outputFilePath)
        {
            throw new NotImplementedException();
        }
    }
}
