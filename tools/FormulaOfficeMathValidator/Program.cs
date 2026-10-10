using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;

const string MathNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/math";
const string Drawing2010Namespace = "http://schemas.microsoft.com/office/drawing/2010/main";

if (args.Length == 0 || args.Contains("--help", StringComparer.OrdinalIgnoreCase))
{
    Console.Error.WriteLine("Usage: FormulaOfficeMathValidator <pptx> [--max-errors 20] [--json <path>]");
    return args.Length == 0 ? 2 : 0;
}

var pptxPath = args[0];
var maxErrors = 20;
string? jsonOutputPath = null;
for (var i = 1; i < args.Length - 1; i++)
{
    if (args[i].Equals("--max-errors", StringComparison.OrdinalIgnoreCase) &&
        int.TryParse(args[i + 1], out var parsed) &&
        parsed > 0)
    {
        maxErrors = parsed;
    }
    else if (args[i].Equals("--json", StringComparison.OrdinalIgnoreCase))
    {
        jsonOutputPath = args[i + 1];
    }
}

void WriteValidatorJson(string json)
{
    Console.WriteLine(json);
    if (!string.IsNullOrEmpty(jsonOutputPath))
    {
        File.WriteAllText(jsonOutputPath, json);
    }
}

if (!File.Exists(pptxPath))
{
    Console.Error.WriteLine($"PPTX not found: {pptxPath}");
    return 2;
}

ValidationResult result;
try
{
    result = ValidatePresentation(pptxPath, maxErrors);
}
catch (Exception ex)
{
    // Emit the same JSON contract instead of a raw stack trace so callers can
    // parse the failure.
    var failure = new ValidationResult(
        Path.GetFullPath(pptxPath),
        0,
        0,
        0,
        1,
        new List<ValidationIssue> { new ValidationIssue($"Package could not be opened or validated: {ex.Message}", string.Empty, string.Empty) });
    Console.WriteLine(JsonSerializer.Serialize(failure, new JsonSerializerOptions { WriteIndented = true }));
    if (!string.IsNullOrEmpty(jsonOutputPath))
    {
        File.WriteAllText(jsonOutputPath, JsonSerializer.Serialize(failure, new JsonSerializerOptions { WriteIndented = true }));
    }
    return 1;
}
var json = JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
WriteValidatorJson(json);
return result.OpenXmlErrorCount == 0 ? 0 : 1;

static ValidationResult ValidatePresentation(string pptxPath, int maxErrors)
{
    using var document = PresentationDocument.Open(pptxPath, false);
    var presentationPart = document.PresentationPart;
    var slideParts = GetSlidePartsInPresentationOrder(presentationPart);
    var a14MathCount = 0;
    var officeMathCount = 0;

    foreach (var slidePart in slideParts)
    {
        var slideXml = slidePart.Slide.OuterXml;
        a14MathCount += CountXmlElement(slideXml, Drawing2010Namespace, "m");
        officeMathCount += CountXmlElement(slideXml, MathNamespace, "oMath");
    }

    // Office2010 is the minimum schema level that covers a14:m math content;
    // the parameterless Office2007 default would skip the very markup this
    // validator exists to gate.
    var validator = new OpenXmlValidator(FileFormatVersions.Office2010);
    var allErrors = validator.Validate(document).ToList();

    // The OpenXml SDK models a14:m (the MS-ODRAWXML math wrapper) as a leaf
    // element, while the MS-ODRAWXML extension spec requires it to contain
    // m:oMath — the exact markup PowerPoint itself writes for shape math.
    // That single complaint is a known SDK limitation: it is skipped instead
    // of failing the gate; every other validation error still does.
    var blocking = new List<ValidationErrorInfo>();
    foreach (var error in allErrors)
    {
        var xpath = error.Path?.XPath ?? string.Empty;
        var description = error.Description ?? string.Empty;
        if (xpath.Contains("a14:m", StringComparison.Ordinal) && description.Contains("leaf element", StringComparison.Ordinal))
        {
            continue;
        }
        blocking.Add(error);
    }

    var errors = blocking
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
        blocking.Count,
        errors);
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
    IReadOnlyList<ValidationIssue> OpenXmlErrors);

public sealed record ValidationIssue(string Description, string XPath, string PartUri);
