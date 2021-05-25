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
            // Issue 01
            // When converting List Paragrapg from this document, the Indentation 
            // is not preserved.
            ConvertToPdf("CohortReport-417b95f2-8c4c-46c9-9b50-f6060c9f89ba.docx");
            // Issue 01-a
            // But this document, which appear to be the same converts correctly
            // What is it that Spire PDF Writer is interpretting different please?
            ConvertToPdf("CohortReport.docx");

            // Issue 02 - This image has a large graphic in it, it stretches and distorts
            ConvertToPdf("TestBigGraphic.docx");
            
           
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
