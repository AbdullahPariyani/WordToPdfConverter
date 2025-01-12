using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Drawing;
using System.IO;

class Program
{
    static void Main(string[] args)
    {
        string filePath = "input_ar.docx";
        string imagePath = "qr.jpg";

        // Create a new document or open an existing one
        using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(filePath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            // Add a main document part
            MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            // Add the image to the document
            AddImageToDocument(mainPart, imagePath);
        }
    }

    private static void AddImageToDocument(MainDocumentPart mainPart, string imagePath)
    {
        // Add an image part to the document
        ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

        // Load the image into the image part
        using (FileStream stream = new FileStream(imagePath, FileMode.Open))
        {
            imagePart.FeedData(stream);
        }

        // Get the ID of the image part
        string imagePartId = mainPart.GetIdOfPart(imagePart);

        // Add the image to the document body
        AddImageToBody(mainPart, imagePartId);
    }

    private static void AddImageToBody(MainDocumentPart mainPart, string relationshipId)
    {
        // Create a new drawing element
        var element = new Drawing(
            new Inline(
                new Extent() { Cx = 990000L, Cy = 792000L }, // Size in EMUs (1 inch = 914400 EMUs)
                new EffectExtent()
                {
                    LeftEdge = 0L,
                    TopEdge = 0L,
                    RightEdge = 0L,
                    BottomEdge = 0L
                },
                new DocProperties()
                {
                    Id = (UInt32Value)1U,
                    Name = "Picture"
                },
                new NonVisualGraphicFrameDrawingProperties(new GraphicFrameLocks() { NoChangeAspect = true }),
                new Graphic(
                    new GraphicData(
                        new Picture(
                            new NonVisualPictureProperties(
                                new NonVisualDrawingProperties()
                                {
                                    Id = (UInt32Value)0U,
                                    Name = "New Bitmap Image.jpg"
                                },
                                new NonVisualPictureDrawingProperties()),
                            new BlipFill(
                                new Blip() { Embed = relationshipId },
                                new Stretch(new FillRectangle())),
                            new ShapeProperties(
                                new Transform2D(
                                    new Offset() { X = 0L, Y = 0L },
                                    new Extents() { Cx = 990000L, Cy = 792000L }),
                                new PresetGeometry(new AdjustValueList())
                                { Preset = ShapeTypeValues.Rectangle })))
                    )
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
            )
            {
                DistanceFromTop = (UInt32Value)0U,
                DistanceFromBottom = (UInt32Value)0U,
                DistanceFromLeft = (UInt32Value)0U,
                DistanceFromRight = (UInt32Value)0U,
                EditId = "50D07946"
            });

        // Append the drawing element to the document body
        mainPart.Document.Body.AppendChild(new Paragraph(new Run(element)));
    }
}
