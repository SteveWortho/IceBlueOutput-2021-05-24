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
            ConvertToPdf(@"\Bug991\AssessmentFeedbackOriginal.docx");

            // If the images are placed on seperate pages or single pages it converts fine
            // so it must be the table cell set to back that is causing the issue

        }

        public static void ConvertToPdf(string sourceDocx)
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
    }
}
