using System;
using System.Drawing;
using Spire.Doc;
using Spire.Doc.Fields;
using Spire.Doc.Documents; // Import necessary namespaces for Paragraph, BuiltinStyle, and TextWrappingStyle

namespace InsertImageCore
{
    class Program
    {
        static void Main(string[] args)
        {
            // Define the paths for the input Word document and image
            string documentPath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "input_ar.docx");
            string imagePath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "qr.jpg");
            
            // Load the Word document
            Document doc = new Document();
            doc.LoadFromFile(documentPath);

            // Get the first section and first paragraph of the document
            Section section = doc.Sections[0];
            Paragraph paragraph = section.Paragraphs.Count > 0 ? section.Paragraphs[0] : section.AddParagraph();
            paragraph.AppendText("The sample demonstrates how to insert an image into a document.");
            paragraph.ApplyStyle(BuiltinStyle.Heading2);
            
            // Add a second paragraph
            paragraph = section.AddParagraph();
            paragraph.AppendText("The above is a picture.");

            // Load and manipulate the image
            Bitmap p = new Bitmap(Image.FromFile(imagePath));  // Load image from file

            // Rotate image and insert into the document
            p.RotateFlip(RotateFlipType.Rotate90FlipX);  // Rotate the image

            // Create a picture object for the Word document
            DocPicture picture = new DocPicture(doc);
            picture.LoadImage(p);

            // Set the picture's position and size
            picture.HorizontalPosition = 50.0F;
            picture.VerticalPosition = 60.0F;
            picture.Width = 200;
            picture.Height = 200;

            // Set text wrapping style
            picture.TextWrappingStyle = TextWrappingStyle.Through;
            
            // Insert the picture into the second paragraph
            paragraph.ChildObjects.Insert(0, picture);

            // Save the document to a new file
            string output = "InsertImageAtSpecifiedLocation.docx";
            doc.SaveToFile(output, FileFormat.Docx);
            
            // Open the saved document
            Viewer(output);
        }

        // Helper method to open the generated document
        private static void Viewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);  // Launch the document
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
