using org.apache.pdfbox.pdmodel;
using org.apache.pdfbox.util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using org.apache.pdfbox.pdmodel;
using org.apache.pdfbox.util;
using org.apache.pdfbox;
using java.io;

namespace GetTextFromPDFToExcel
{
    public partial class Program
    {
        private static string GetTextPDF(string pdfFile)
        {
            PDDocument doc = null;
            try
            {
                doc = PDDocument.load(pdfFile);
              PDFTextStripper stripper = new PDFTextStripper();
                return stripper.getText(doc);
            }
            finally
            {
                if (doc != null)
                {
                    doc.close();
                }
            }
        }
    }
}
