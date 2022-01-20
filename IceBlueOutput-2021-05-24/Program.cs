using Spire.Doc;
using Spire.Doc.Documents;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IceBlueOutput_2021_05_24
{
    class Program
    {

        public const string AppPath = @"C:\Users\SteveWorthington\source\repos\IceBlueOutput-2021-05-24\IceBlueOutput-2021-05-24\";

        static void Main(string[] args)
        {
            // Bug #991 - Convert to PDF
            // Fails when all the images are in a single document set to
            // behind text and fixed position on page
            //ConvertToPdf(@"\Bug991\AssessmentFeedbackOriginal.docx");

            // Bug #1010 - Convert to PDF gets in a loop of death
            // Fails on latest version 9.8.6.0 of Spire
            // Works on previous version 7.9.5.4046?
            //ConvertToPdf(@"\Bug1010\Construction_Risk_v2__20211026 2220.docx");
            // If the images are placed on seperate pages or single pages it converts fine
            // so it must be the table cell set to back that is causing the issue

            // 20-Jan-2022 Bug #1050
            // Guage image does not appear in PDF but does in Word
            ConvertToPdf(@"\Bug1050\Temp-124217d2-2032-4384-812e-6be0a41ca600.docx");

        }

        public static void ConvertToPdf(string sourceDocx)
        {
            try
            {
                var template = File.ReadAllBytes($@"{AppPath}\WordDocs\{sourceDocx}");
                var asBytes = template.ToArray();
                string fileName = $@"{AppPath}\PdfDocs\Pdf-{Guid.NewGuid()}.pdf";
                Document document = new Document();

                using (MemoryStream stream = new MemoryStream())
                {
                    stream.Write(asBytes, 0, (int)asBytes.Length);
                    document.LoadFromStream(stream, FileFormat.Docx, XHTMLValidationType.Transitional);
                }
                document.SaveToFile(fileName, FileFormat.PDF);
                System.Diagnostics.Process.Start(fileName);
            }
            catch (Exception exception)
            {
                var error = exception.Message;
                throw;
            }

        }
    }
}
