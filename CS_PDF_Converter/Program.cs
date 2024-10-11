//using DocumentFormat.OpenXml.Packaging;
//using PdfSharpCore.Drawing;
//using PdfSharpCore.Pdf;

//class Program
//{
//    static void Main()
//    {
//        string docxPath = "D:\\MyDemos\\netcodedocng\\Core_API_DocOpener\\Core_API_DocOpener\\Files\\File1.docx";
//        string pdfPath = "D:\\MyDemos\\netcodedocng\\Core_API_DocOpener\\Core_API_DocOpener\\Files\\File1.pdf";

//        try
//        {
//            // Open the DOCX file
//            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(docxPath, false))
//            {
//                // Create a new PDF document
//                using (PdfDocument pdfDoc = new PdfDocument())
//                {
//                    PdfPage pdfPage = pdfDoc.AddPage();
//                    XGraphics gfx = XGraphics.FromPdfPage(pdfPage);
//                    XFont font = new XFont("Verdana", 12, XFontStyle.Regular);

//                    // Loop through each paragraph in the Word document
//                    var paragraphs = wordDoc.MainDocumentPart.Document.Body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>();
//                    int yPoint = 0;
//                    foreach (var paragraph in paragraphs)
//                    {
//                        gfx.DrawString(paragraph.InnerText, font, XBrushes.Black, new XRect(0, yPoint, pdfPage.Width, pdfPage.Height), XStringFormats.TopLeft);
//                        yPoint += 20; // Move to the next line
//                    }

//                    // Save the PDF document
//                    pdfDoc.Save(pdfPath);
//                }
//            }

//            Console.WriteLine("DOCX file has been converted to PDF successfully.");
//        }
//        catch (Exception ex)
//        {
//            Console.WriteLine("Error: " + ex.Message);
//        }
//    }
//}

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf;
using System;
using System.IO;
using System.Linq;

class Program
{
    static void Main()
    {
        string docxPath = "D:\\MyDemos\\netcodedocng\\Core_API_DocOpener\\Core_API_DocOpener\\Files\\File1.docx";
        string pdfPath = "D:\\MyDemos\\netcodedocng\\Core_API_DocOpener\\Core_API_DocOpener\\Files\\File1.pdf";

        try
        {
            // Open the DOCX file
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(docxPath, false))
            {


                var mainPart = wordDoc.MainDocumentPart;
               
                var imageParts = mainPart.ImageParts;

                

                // Create a new PDF document
                using (PdfDocument pdfDoc = new PdfDocument())
                {
                    PdfPage pdfPage = pdfDoc.AddPage();
                    XGraphics gfx = XGraphics.FromPdfPage(pdfPage);
                    XFont font = new XFont("Verdana", 12, XFontStyle.Regular);

                    // Loop through each element in the Word document
                    var bodyElements = wordDoc.MainDocumentPart.Document.Body.Elements();
                    // var bodyElements = wordDoc.MainDocumentPart.Document.Body.ChildElements;
                    int yPoint = 0;
                    foreach (var element in bodyElements)
                    {
                        if (element is Paragraph paragraph)
                        {
                            gfx.DrawString(paragraph.InnerText, font, XBrushes.Black, new XRect(0, yPoint, pdfPage.Width, pdfPage.Height), XStringFormats.TopLeft);
                            yPoint += 20; // Move to the next line
                        }
                        else if (element is DocumentFormat.OpenXml.Wordprocessing.Drawing drawing)
                        {
                            var blip = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
                            if (blip != null)
                            {
                                var imagePart = (ImagePart)wordDoc.MainDocumentPart.GetPartById(blip.Embed.Value);
                                using (Stream? stream = imagePart.GetStream())
                                {
                                    var image = XImage.FromStream(() => stream);
                                    gfx.DrawImage(image, 0, yPoint);
                                    yPoint += (int)image.PointHeight; // Move to the next line after the image
                                }
                            }
                        }
                    }

                    // Save the PDF document
                    pdfDoc.Save(pdfPath);
                }
            }

            Console.WriteLine("DOCX file has been converted to PDF successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
