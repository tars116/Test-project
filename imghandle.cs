public class OpenXmlWordDocumentGenerator : IWordDocumentGenerator
{
    public void GenerateDocument(Dictionary<string, string> contentDictionary, string outputPath)
    {
        using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document))
        {
            MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document();
            Body body = new Body();
            mainPart.Document.Append(body);

            foreach (var entry in contentDictionary)
            {
                Paragraph paragraph = new Paragraph(new Run(new Text($"{entry.Key}:")));
                body.Append(paragraph);

                // Check if the value contains HTML
                if (IsHtml(entry.Value))
                {
                    // Convert HTML to Word format
                    ParseHtmlContent(entry.Value, mainPart, body);
                }
                else
                {
                    // Add plain text
                    Run run = new Run(new Text(entry.Value));
                    paragraph.Append(run);
                }
            }

            mainPart.Document.Save();
        }
    }

    // Check if a string contains HTML
    private bool IsHtml(string input)
    {
        return input.Contains("<") && input.Contains(">");
    }

    // Parse HTML content and convert it to Word format
    private void ParseHtmlContent(string htmlContent, MainDocumentPart mainPart, Body body)
    {
        // Use HtmlAgilityPack or a similar HTML parser to handle the HTML
        var doc = new HtmlDocument();
        doc.LoadHtml(htmlContent);

        // Iterate through all the nodes in the HTML content
        foreach (var node in doc.DocumentNode.ChildNodes)
        {
            if (node.Name == "p" || node.Name == "div")
            {
                // For paragraph or div, create a new paragraph in the Word document
                Paragraph paragraph = new Paragraph(new Run(new Text(node.InnerText)));
                body.Append(paragraph);
            }
            else if (node.Name == "img")
            {
                // Handle image tag
                string imageUrl = node.GetAttributeValue("src", null);
                if (imageUrl != null)
                {
                    AddImageToWord(mainPart, body, imageUrl);
                }
            }
            // You can extend this to handle other HTML tags (bold, italic, links, etc.)
        }
    }

    // Add an image to the Word document
    private void AddImageToWord(MainDocumentPart mainPart, Body body, string imageUrl)
    {
        // Image processing
        ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);  // Change type as needed

        // Assuming the imageUrl is a file path or you can load the image as a byte array
        byte[] imageBytes = File.ReadAllBytes(imageUrl);
        using (var stream = new MemoryStream(imageBytes))
        {
            imagePart.FeedData(stream);
        }

        var imageId = mainPart.GetIdOfPart(imagePart);

        // Add image to document (similar to your previous implementation)
        var element =
             new DW.Inline(
                 new DW.Extent() { Cx = 990000L, Cy = 792000L },
                 new DW.EffectExtent()
                 {
                     LeftEdge = 0L,
                     TopEdge = 0L,
                     RightEdge = 0L,
                     BottomEdge = 0L
                 },
                 new DW.DocProperties()
                 {
                     Id = (UInt32Value)1U,
                     Name = "Picture"
                 },
                 new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoChangeAspect = true }),
                 new A.Graphic(
                     new A.GraphicData(
                         new PIC.Picture(
                             new PIC.NonVisualPictureProperties(
                                 new PIC.NonVisualDrawingProperties()
                                 {
                                     Id = (UInt32Value)0U,
                                     Name = "New Image.jpg"
                                 },
                                 new PIC.NonVisualPictureDrawingProperties()),
                             new PIC.BlipFill(
                                 new A.Blip(
                                     new A.BlipExtensionList(
                                         new A.BlipExtension()
                                         {
                                             Uri =
                                             "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                         }))
                                 {
                                     Embed = imageId,
                                     CompressionState = A.BlipCompressionValues.Print
                                 },
                                 new A.Stretch(new A.FillRectangle())),
                             new PIC.ShapeProperties(
                                 new A.Transform2D(
                                     new A.Offset() { X = 0L, Y = 0L },
                                     new A.Extents() { Cx = 990000L, Cy = 792000L }),
                                 new A.PresetGeometry(new A.AdjustValueList())
                                 { Preset = A.ShapeTypeValues.Rectangle })))
                     { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
             {
                 DistanceFromTop = (UInt32Value)0U,
                 DistanceFromBottom = (UInt32Value)0U,
                 DistanceFromLeft = (UInt32Value)0U,
                 DistanceFromRight = (UInt32Value)0U,
             };

        var drawingElement = new Drawing(element);
        var paragraphWithImage = new Paragraph(new Run(drawingElement));
        body.Append(paragraphWithImage);
    }
}
