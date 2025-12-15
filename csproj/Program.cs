using System;
using System.IO;
using OpenXmlPowerTools;
using DocumentFormat.OpenXml.Packaging;

class Program
{
    static void Main(string[] args)
    {
        // Check for mode flag
        if (args.Length > 0 && args[0] == "--sheets")
        {
            CompareSpreadsheets(args);
        }
        else if (args.Length > 0 && args[0] == "--presentations")
        {
            ComparePresentations(args);
        }
        else
        {
            CompareDocuments(args);
        }
    }

    static void CompareDocuments(string[] args)
    {
        if (args.Length != 4)
        {
            Console.WriteLine("Usage: redlines <author_tag> <original_path.docx> <modified_path.docx> <redline_path.docx>");
            return;
        }

        string authorTag = args[0];
        string originalFilePath = args[1];
        string modifiedFilePath = args[2];
        string outputFilePath = args[3];

        if (!File.Exists(originalFilePath) || !File.Exists(modifiedFilePath))
        {
            Console.WriteLine("Error: One or both files do not exist.");
            return;
        }

        try
        {
            var originalBytes = File.ReadAllBytes(originalFilePath);
            var modifiedBytes = File.ReadAllBytes(modifiedFilePath);
            var originalDocument = new WmlDocument(originalFilePath, originalBytes);
            var modifiedDocument = new WmlDocument(modifiedFilePath, modifiedBytes);

            var comparisonSettings = new WmlComparerSettings
            {
                AuthorForRevisions = authorTag,
                DetailThreshold = 0,
                TrackFormattingChanges = true,
                LogCallback = Console.WriteLine,
            };

            var comparisonResults = WmlComparer.Compare(originalDocument, modifiedDocument, comparisonSettings);
            var revisions = WmlComparer.GetRevisions(comparisonResults, comparisonSettings);

            // Output results
            Console.WriteLine($"Revisions found: {revisions.Count}");

            File.WriteAllBytes(outputFilePath, comparisonResults.DocumentByteArray);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
            Console.WriteLine("Detailed Stack Trace:");
            Console.WriteLine(ex.StackTrace);
        }
    }

    static void CompareSpreadsheets(string[] args)
    {
        if (args.Length != 5)
        {
            Console.WriteLine("Usage: redlines --sheets <author_tag> <original_path.xlsx> <modified_path.xlsx> <redline_path.xlsx>");
            return;
        }

        string authorTag = args[1];
        string originalFilePath = args[2];
        string modifiedFilePath = args[3];
        string outputFilePath = args[4];

        if (!File.Exists(originalFilePath) || !File.Exists(modifiedFilePath))
        {
            Console.WriteLine("Error: One or both files do not exist.");
            return;
        }

        try
        {
            var originalBytes = File.ReadAllBytes(originalFilePath);
            var modifiedBytes = File.ReadAllBytes(modifiedFilePath);
            var originalDocument = new SmlDocument(originalFilePath, originalBytes);
            var modifiedDocument = new SmlDocument(modifiedFilePath, modifiedBytes);

            var comparisonSettings = new SmlComparerSettings
            {
                AuthorForChanges = authorTag,
                CompareValues = true,
                CompareFormulas = true,
                CompareFormatting = true,
                CompareSheetStructure = true,
                LogCallback = Console.WriteLine,
            };

            var markedWorkbook = SmlComparer.ProduceMarkedWorkbook(originalDocument, modifiedDocument, comparisonSettings);

            // Output results
            Console.WriteLine($"Spreadsheet comparison complete");

            File.WriteAllBytes(outputFilePath, markedWorkbook.DocumentByteArray);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
            Console.WriteLine("Detailed Stack Trace:");
            Console.WriteLine(ex.StackTrace);
        }
    }

    static void ComparePresentations(string[] args)
    {
        if (args.Length != 5)
        {
            Console.WriteLine("Usage: redlines --presentations <author_tag> <original_path.pptx> <modified_path.pptx> <redline_path.pptx>");
            return;
        }

        string authorTag = args[1];
        string originalFilePath = args[2];
        string modifiedFilePath = args[3];
        string outputFilePath = args[4];

        if (!File.Exists(originalFilePath) || !File.Exists(modifiedFilePath))
        {
            Console.WriteLine("Error: One or both files do not exist.");
            return;
        }

        try
        {
            var originalBytes = File.ReadAllBytes(originalFilePath);
            var modifiedBytes = File.ReadAllBytes(modifiedFilePath);
            var originalDocument = new PmlDocument(originalFilePath, originalBytes);
            var modifiedDocument = new PmlDocument(modifiedFilePath, modifiedBytes);

            var comparisonSettings = new PmlComparerSettings
            {
                AuthorForChanges = authorTag,
                CompareTextContent = true,
                CompareShapeStructure = true,
                CompareSlideStructure = true,
                CompareImageContent = true,
                AddSummarySlide = true,
                AddNotesAnnotations = true,
                LogCallback = Console.WriteLine,
            };

            var markedPresentation = PmlComparer.ProduceMarkedPresentation(originalDocument, modifiedDocument, comparisonSettings);

            // Output results
            Console.WriteLine($"Presentation comparison complete");

            File.WriteAllBytes(outputFilePath, markedPresentation.DocumentByteArray);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
            Console.WriteLine("Detailed Stack Trace:");
            Console.WriteLine(ex.StackTrace);
        }
    }
}

