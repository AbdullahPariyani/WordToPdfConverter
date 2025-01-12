using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Drawing;
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
                { "{{amountAdmin}}", "30" },
                { "{{amountTotal}}", "2896.33" },
                { "{{periodFrom}}", "14/05/2024" },
                { "{{periodTo}}", "13/05/2025" },
                { "{{image}}", "qr.jpg" } // Image file path
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
                // Get the body of the document
                var body = wordDoc.MainDocumentPart.Document.Body;

                // Iterate through all text elements in the document
                foreach (var text in body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>())
                {
                    // Check each placeholder in the dictionary
                    foreach (var replacement in replacements)
                    {
                        if (text.Text.Contains(replacement.Key))
                        {
                            if (replacement.Key == "{{image}}")
                            {
                                Console.WriteLine("Image Found: " + text);
                                // Replace the placeholder with an image (for {{image}})
                                InsertImage(wordDoc, text, replacement.Value);
                            }
                            else
                            {
                                // Replace the text placeholder with the corresponding value for other placeholders
                                text.Text = text.Text.Replace(replacement.Key, replacement.Value);
                            }
                        }
                    }
                }

                // Save the changes
                wordDoc.MainDocumentPart.Document.Save();
            }
        }

        // Method to insert an image at the placeholder location in the DOCX
        static void InsertImage(WordprocessingDocument wordDoc, DocumentFormat.OpenXml.Wordprocessing.Text textElement, string imagePath)
        {


            // Create a new image part in the document's main part
            ImagePart imagePart = wordDoc.MainDocumentPart.AddImagePart(ImagePartType.Jpeg);

            // Copy the image file into the image part
            using (FileStream fileStream = new FileStream(imagePath, FileMode.Open))
            {
                imagePart.FeedData(fileStream);
            }

            // Get the relationship ID for the image part
            string relationshipId = wordDoc.MainDocumentPart.GetIdOfPart(imagePart);

            // Create a new run that will contain the image
            var run = new DocumentFormat.OpenXml.Wordprocessing.Run();
            var runProperties = new DocumentFormat.OpenXml.Wordprocessing.RunProperties();
            run.Append(runProperties);

            // Create a picture element for the image
            var picture = new DocumentFormat.OpenXml.Drawing.Picture(
                new DocumentFormat.OpenXml.Drawing.NonVisualPictureProperties(
                    new DocumentFormat.OpenXml.Drawing.NonVisualDrawingProperties() { Id = 1, Name = "QR Image" },
                    new DocumentFormat.OpenXml.Drawing.NonVisualPictureDrawingProperties()),
                new DocumentFormat.OpenXml.Drawing.BlipFill(
                    new DocumentFormat.OpenXml.Drawing.Blip() { Embed = relationshipId }),
                new DocumentFormat.OpenXml.Drawing.ShapeProperties());

            // Append the picture to the run
            run.Append(picture);

            // Ensure textElement is not null and has a valid Parent (which should be a Paragraph)
            var paragraph = textElement?.Parent as DocumentFormat.OpenXml.Wordprocessing.Paragraph;

            // Check if paragraph is still null, which could happen if the parent isn't a Paragraph
            if (paragraph == null)
            {
                throw new InvalidOperationException("Text element must be inside a paragraph.");
            }
            // Insert the run (with image) into the paragraph at the position of the text element
            paragraph.InsertAfter(run, textElement);

            // Remove the original text element (which was a placeholder)
            textElement.Remove();
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
