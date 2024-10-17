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

                // Check if the value contains HTML
                if (IsHtml(entry.Value))
                {
                    // Convert the HTML content into Word OpenXML elements
                    ConvertHtmlToOpenXml(mainPart, entry.Value, body);
                }
                else
                {
                    // Add plain text
                    Run run = new Run(new Text(entry.Value));
                    paragraph.Append(run);
                }

                body.Append(paragraph);
            }

            mainPart.Document.Save();
        }
    }

    // Convert HTML content to Word OpenXML elements, including images
    private void ConvertHtmlToOpenXml(MainDocumentPart mainPart, string htmlContent, Body body)
    {
        var htmlDoc = new HtmlAgilityPack.HtmlDocument();
        htmlDoc.LoadHtml(htmlContent);

        foreach (var node in htmlDoc.DocumentNode.ChildNodes)
        {
            if (node.Name == "img")
            {
                // Process image tags
                var imageSrc = node.GetAttributeValue("src", null);
                if (!string.IsNullOrEmpty(imageSrc))
                {
                    // Add image to the Word document
                    AddImageToDocument(mainPart, imageSrc, body);
                }
            }
            else
            {
                // Process other HTML elements, e.g., <p>, <b>, <i>, <a>, etc.
                Run run = new Run(new Text(node.InnerText));
                Paragraph paragraph = new Paragraph(run);
                body.Append(paragraph);
            }
        }
    }

    // Add image to Word document from HTML <img> tag
    private void AddImageToDocument(MainDocumentPart mainPart, string imageUrl, Body body)
    {
        // Logic to download or get the image bytes from imageUrl and insert into Word
        var imagePart = mainPart.AddImagePart(ImagePartType.Jpeg); // Assuming JPEG format

        using (var webClient = new WebClient())
        {
            byte[] imageBytes = webClient.DownloadData(imageUrl);
            using (var stream = new MemoryStream(imageBytes))
            {
                imagePart.FeedData(stream);
            }
        }

        AddImageToBody(mainPart.GetIdOfPart(imagePart), body);
    }

    // Helper method to add image to the body
    private void AddImageToBody(string relationshipId, Body body)
    {
        var element =
             new Drawing(
                 new DW.Inline(
                     new DW.Extent() { Cx = 990000L, Cy = 792000L },
                     new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                     new DW.DocProperties() { Id = (UInt32Value)1U, Name = "Picture" },
                     new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoChangeAspect = true }),
                     new A.Graphic(
                         new A.GraphicData(
                             new PIC.Picture(
                                 new PIC.NonVisualPictureProperties(
                                     new PIC.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "New Bitmap Image.jpg" },
                                     new PIC.NonVisualPictureDrawingProperties()),
                                 new PIC.BlipFill(
                                     new A.Blip() { Embed = relationshipId },
                                     new A.Stretch(new A.FillRectangle())),
                                 new PIC.ShapeProperties(
                                     new A.Transform2D(
                                         new A.Offset() { X = 0L, Y = 0L },
                                         new A.Extents() { Cx = 990000L, Cy = 792000L }),
                                     new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
                         ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
             ) { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U });

        body.AppendChild(new Paragraph(new Run(element)));
    }

    // Check if a string contains HTML
    private bool IsHtml(string input)
    {
        return input.Contains("<") && input.Contains(">");
    }
}
