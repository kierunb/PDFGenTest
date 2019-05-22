using NetOffice.WordApi;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NetOffice;
using Word = NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using NetOffice.WordApi.Tools.Contribution;

namespace PDFGenTest.Converters
{
    public class NetOfficeWord : IConvert
    {
        public void ConvertToPDF(string inputFilePath, string outputFilePath)
        {

            using (Word.Application wordApp = new Word.Application())
            {
                wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                var doc = wordApp.Documents.Open(inputFilePath);
                doc.SaveAs(outputFilePath, fileFormat: WdExportFormat.wdExportFormatPDF);
                doc.TablesOfContents[0]?.Update();
                wordApp.Quit(saveChanges: false);
                wordApp.Dispose();
            }

        }

        public void ConvertToPDF(string[] inputFilePath, string[] outputFilePath)
        {
            throw new NotImplementedException();
        }
    }
}
