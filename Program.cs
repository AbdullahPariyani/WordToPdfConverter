using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Collections.Generic;

namespace DocxToPdfConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to the input and output files
            string docxFilePath = "input_ar.docx";  // Adjust the path accordingly
            string copiedDocxFilePath = "input_copy.docx"; // New copy for modifications
            string pdfFilePath = "output.pdf";   // Adjust the path accordingly

            // Dictionary of placeholders and their replacements
            var replacements = new Dictionary<string, string>
            {
                { "{{invoiceNameE}}", "ABDULLAH MOHAMMAD ILIYAS PARIYANI" },
                { "{{invoiceNameA}}", "عبدالله محمد الياس بارياني" },
                { "{{invoiceNumber}}", "2024/DNU65-700/5766" },
                { "{{invoiceDate}}", "13-05-2024 00:00:00" },
                { "{{policyNumber}}", "330519770" },
                { "{{amount}}", "2488.55" },
                { "{{amountAdmin}}", "500" },
                { "{{amountTotal}}", "2896.33" },
                { "{{periodFrom}}", "14/05/2024" },
                { "{{periodTo}}", "13/05/2025" },
                { "{{abdullah}}", "13/05/2025" },
                {"{{image}}",""}
            };

            try
            {
                // Step 1: Create a copy of the original DOCX file
                CreateDocxCopy(docxFilePath, copiedDocxFilePath);
                Console.WriteLine("Created a copy of the original document.");

                // Step 2: Replace text in the copied DOCX file
                ReplaceTextInDocx(copiedDocxFilePath, replacements);
                Console.WriteLine("Text replaced successfully.");

                // Step 3: Convert the modified copied DOCX to PDF using LibreOffice
                ConvertDocxToPdf(copiedDocxFilePath, pdfFilePath);
                Console.WriteLine("Document converted to PDF successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Method to create a copy of the original DOCX file
        static void CreateDocxCopy(string originalFilePath, string copyFilePath)
        {
            // Copy the original file to a new location (to avoid modifying the original)
            File.Copy(originalFilePath, copyFilePath, true); // `true` to overwrite if exists
        }

        // Method to replace text in the DOCX file
        static void ReplaceTextInDocx(string docxFilePath, Dictionary<string, string> replacements)
        {
            // Open the DOCX file
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(docxFilePath, true))
            {
                // Step 1: Replace text in the main document body
                var body = wordDoc.MainDocumentPart.Document.Body;
                ReplaceTextInElement(body, replacements);

                // Step 2: Replace text in headers
                foreach (var headerPart in wordDoc.MainDocumentPart.HeaderParts)
                {
                    ReplaceTextInElement(headerPart.Header, replacements);
                }

                // Step 3: Replace text in footers
                foreach (var footerPart in wordDoc.MainDocumentPart.FooterParts)
                {
                    ReplaceTextInElement(footerPart.Footer, replacements);
                }

                // Save the changes
                wordDoc.MainDocumentPart.Document.Save();
            }
        }

        // Helper method to replace text in a given OpenXml element
        static void ReplaceTextInElement(OpenXmlElement element, Dictionary<string, string> replacements)
        {
            var allTextElements = element.Descendants<Text>();

            // Debug: Check the contents before processing
            foreach (var text in allTextElements)
            {
                Console.WriteLine($"Original text: {text.Text}");  // Debug log

                foreach (var replacement in replacements)
                {
                    if (text.Text.Contains(replacement.Key))
                    {
                        // Replace the placeholder with the corresponding value
                        text.Text = text.Text.Replace(replacement.Key, replacement.Value);
                    }
                }
            }
        }

        // Method to convert DOCX to PDF using LibreOffice
        static void ConvertDocxToPdf(string docxFilePath, string pdfFilePath)
        {
            // Path to the LibreOffice executable (adjust according to your installation)
            string libreOfficePath = @"C:\Program Files\LibreOffice\program\soffice.exe"; // Modify as needed

            // Construct the command to convert DOCX to PDF
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = libreOfficePath,
                Arguments = $"--headless --convert-to pdf --outdir \"{System.IO.Path.GetDirectoryName(pdfFilePath)}\" \"{docxFilePath}\"",
                CreateNoWindow = true,
                UseShellExecute = false
            };

            // Start the process to convert the file
            Process process = Process.Start(startInfo);
            process.WaitForExit(); // Wait for the process to finish

            // Check if the PDF was created
            if (System.IO.File.Exists(pdfFilePath))
            {
                Console.WriteLine("PDF created successfully.");
            }
            else
            {
                Console.WriteLine("Error: PDF conversion failed.");
            }
        }
    }
}
