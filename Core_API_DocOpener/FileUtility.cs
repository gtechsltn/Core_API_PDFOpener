using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.Office.Interop.Word;

namespace Core_API_DocOpener
{
    /// <summary>
    /// Class to convert Word File to PDF
    /// </summary>
    public class FileUtility
    {
        /// <summary>
        /// Converts a Word document to a PDF document.
        /// </summary>
        /// <param name="inputFilePath">The path to the input Word document.</param>
        /// <param name="outputFilePath">The path to the output PDF document.</param>
        public void ConvertWordToPdf(string inputFilePath, string outputFilePath)
        {
            if (string.IsNullOrEmpty(inputFilePath))
                throw new ArgumentException("Input file path cannot be null or empty.", nameof(inputFilePath));

            if (string.IsNullOrEmpty(outputFilePath))
                throw new ArgumentException("Output file path cannot be null or empty.", nameof(outputFilePath));

            Application wordApp = new Application();
            Document wordDocument = null;

            try
            {
                wordDocument = wordApp.Documents.Open(inputFilePath);
                wordDocument.ExportAsFixedFormat(outputFilePath, WdExportFormat.wdExportFormatPDF);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Error converting Word document to PDF.", ex);
            }
            finally
            {
                if (wordDocument != null)
                {
                    wordDocument.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDocument);
                }

                wordApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
            }
        }
    }
}
