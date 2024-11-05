public void GenerateFilteredDocument(Dictionary<string, string> contentDic, string outputPath)
{
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document))
    {
        MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
        mainPart.Document = new Document();

        Body body = new Body();
        mainPart.Document.Append(body);

        foreach (var entry in contentDic)
        {
            // Create and format the question paragraph
            if (entry.Key.Contains("question"))
            {
                Paragraph questionParagraph = CreateParagraph($"Question: {entry.Value}");
                body.Append(questionParagraph);
            }

            // Create and format the answer paragraph if the key contains "answer" and it is HTML content
            if (entry.Key.Contains("answers") && IsHtml(entry.Value))
            {
                Paragraph answersParagraph = CreateParagraph("Answers:");
                body.Append(answersParagraph);

                HtmlConverter converter = new HtmlConverter(mainPart);
                converter.ParseHtml(entry.Value);
            }
        }

        mainPart.Document.Save();
    }
}

// Helper method to create a paragraph with alignment and spacing adjustments
private Paragraph CreateParagraph(string text)
{
    Paragraph paragraph = new Paragraph();

    // Set the text in a run
    Run run = new Run(new Text(text));
    paragraph.Append(run);

    // Apply paragraph properties for alignment and spacing
    ParagraphProperties paragraphProperties = new ParagraphProperties();

    // Set spacing to remove extra spacing between paragraphs
    SpacingBetweenLines spacing = new SpacingBetweenLines
    {
        After = "0",  // No spacing after the paragraph
        Before = "0", // No spacing before the paragraph
        Line = "240", // Single line spacing
        LineRule = LineSpacingRuleValues.Auto
    };
    paragraphProperties.Append(spacing);

    // Add alignment (optional: set to justify or any other alignment as needed)
    Justification justification = new Justification { Val = JustificationValues.Left };
    paragraphProperties.Append(justification);

    // Append paragraph properties to the paragraph
    paragraph.PrependChild(paragraphProperties);

    return paragraph;
}

// Helper method to check if content is HTML
private bool IsHtml(string input)
{
    return input.Contains("<") && input.Contains(">");
}

