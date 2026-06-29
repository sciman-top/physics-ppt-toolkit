using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;

const string MathNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/math";
const string Drawing2010Namespace = "http://schemas.microsoft.com/office/drawing/2010/main";

if (args.Length == 0 || args.Contains("--help", StringComparer.OrdinalIgnoreCase))
{
    Console.Error.WriteLine("Usage: FormulaOfficeMathValidator <pptx> [--max-errors 20]");
    return args.Length == 0 ? 2 : 0;
}

var pptxPath = args[0];
var maxErrors = 20;
for (var i = 1; i < args.Length - 1; i++)
{
    if (args[i].Equals("--max-errors", StringComparison.OrdinalIgnoreCase) &&
        int.TryParse(args[i + 1], out var parsed) &&
        parsed > 0)
    {
        maxErrors = parsed;
    }
}

if (!File.Exists(pptxPath))
{
    Console.Error.WriteLine($"PPTX not found: {pptxPath}");
    return 2;
}

var result = ValidatePresentation(pptxPath, maxErrors);
var json = JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
Console.WriteLine(json);
return result.OpenXmlErrorCount == 0 ? 0 : 1;

static ValidationResult ValidatePresentation(string pptxPath, int maxErrors)
{
    using var document = PresentationDocument.Open(pptxPath, false);
    var presentationPart = document.PresentationPart;
    var slideParts = GetSlidePartsInPresentationOrder(presentationPart);
    var a14MathCount = 0;
    var officeMathCount = 0;
    var slideMathCounts = new List<SlideMathCount>();

    foreach (var slidePart in slideParts)
    {
        var slide = slidePart.Slide;
        var slideXml = slide.OuterXml;
        var slideNumber = slideMathCounts.Count + 1;
        var slideA14 = CountXmlElement(slideXml, Drawing2010Namespace, "m");
        var slideOmml = CountXmlElement(slideXml, MathNamespace, "oMath");
        a14MathCount += slideA14;
        officeMathCount += slideOmml;
        if (slideA14 > 0 || slideOmml > 0)
        {
            slideMathCounts.Add(new SlideMathCount(slideNumber, slideA14, slideOmml));
        }
        else
        {
            slideMathCounts.Add(new SlideMathCount(slideNumber, 0, 0));
        }
    }

    var validator = new OpenXmlValidator();
    var errors = validator.Validate(document)
        .Take(maxErrors)
        .Select(e => new ValidationIssue(
            e.Description ?? string.Empty,
            e.Path?.XPath ?? string.Empty,
            e.Part?.Uri.ToString() ?? string.Empty))
        .ToList();

    return new ValidationResult(
        Path.GetFullPath(pptxPath),
        slideParts.Count,
        a14MathCount,
        officeMathCount,
        errors.Count,
        errors,
        slideMathCounts.Where(s => s.A14MathCount > 0 || s.OfficeMathCount > 0).ToList());
}

static List<SlidePart> GetSlidePartsInPresentationOrder(PresentationPart? presentationPart)
{
    if (presentationPart?.Presentation?.SlideIdList is null)
    {
        return [];
    }

    var parts = new List<SlidePart>();
    foreach (var slideId in presentationPart.Presentation.SlideIdList.Elements<SlideId>())
    {
        var relId = slideId.RelationshipId?.Value;
        if (string.IsNullOrWhiteSpace(relId))
        {
            continue;
        }

        if (presentationPart.GetPartById(relId) is SlidePart slidePart)
        {
            parts.Add(slidePart);
        }
    }

    return parts;
}

static int CountXmlElement(string xml, string namespaceUri, string localName)
{
    using var reader = System.Xml.XmlReader.Create(new StringReader(xml));
    var count = 0;
    while (reader.Read())
    {
        if (reader.NodeType == System.Xml.XmlNodeType.Element &&
            reader.LocalName == localName &&
            reader.NamespaceURI == namespaceUri)
        {
            count++;
        }
    }

    return count;
}

public sealed record ValidationResult(
    string PptxPath,
    int SlideCount,
    int A14MathCount,
    int OfficeMathCount,
    int OpenXmlErrorCount,
    IReadOnlyList<ValidationIssue> OpenXmlErrors,
    IReadOnlyList<SlideMathCount> SlidesWithMath);

public sealed record ValidationIssue(string Description, string XPath, string PartUri);

public sealed record SlideMathCount(int SlideNumber, int A14MathCount, int OfficeMathCount);
