using System.IO.Compression;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Xml;

bool wantSummary = args.Contains("--summary");
string[] folders = args.Where(a => a != "--summary").ToArray();

if (folders.Length < 2)
{
    Console.WriteLine("Usage: ExcelCompare <folder1> <folder2> [--summary]");
    Console.WriteLine("Compares all matching .xlsx files recursively and reports cell value differences.");
    Console.WriteLine("  --summary   Send the diff report to Claude AI for a plain-language summary.");
    Console.WriteLine("              Requires the ANTHROPIC_API_KEY environment variable to be set.");
    return 1;
}

string folder1 = folders[0].TrimEnd('\\', '/');
string folder2 = folders[1].TrimEnd('\\', '/');

if (!Directory.Exists(folder1)) { Console.WriteLine($"Folder not found: {folder1}"); return 1; }
if (!Directory.Exists(folder2)) { Console.WriteLine($"Folder not found: {folder2}"); return 1; }

int identical = 0, different = 0, missing = 0;
var report = new List<string>();

// Compare every .xlsx in folder1 against the matching file in folder2
foreach (string file1 in Directory.GetFiles(folder1, "*.xlsx", SearchOption.AllDirectories))
{
    string relative = file1[(folder1.Length + 1)..];
    string file2 = Path.Combine(folder2, relative);

    if (!File.Exists(file2))
    {
        missing++;
        report.Add($"MISSING IN FOLDER2: {relative}");
        continue;
    }

    try
    {
        var diffs = CompareExcelFiles(file1, file2);
        if (diffs.Count == 0)
        {
            identical++;
        }
        else
        {
            different++;
            report.Add($"DIFFER ({diffs.Count} cell(s)): {relative}");
            foreach (string d in diffs.Take(5))
                report.Add($"  {d}");
            if (diffs.Count > 5)
                report.Add($"  ... and {diffs.Count - 5} more");
        }
    }
    catch (Exception ex)
    {
        report.Add($"ERROR reading {relative}: {ex.Message}");
    }
}

// Find files that exist in folder2 but not in folder1
foreach (string file2 in Directory.GetFiles(folder2, "*.xlsx", SearchOption.AllDirectories))
{
    string relative = file2[(folder2.Length + 1)..];
    if (!File.Exists(Path.Combine(folder1, relative)))
    {
        missing++;
        report.Add($"MISSING IN FOLDER1: {relative}");
    }
}

Console.WriteLine("=== SUMMARY ===");
Console.WriteLine($"Identical:     {identical}");
Console.WriteLine($"Different:     {different}");
Console.WriteLine($"Missing/extra: {missing}");
Console.WriteLine($"Total checked: {identical + different}");

if (report.Count > 0)
{
    Console.WriteLine("\n=== DETAILS ===");
    foreach (string line in report)
        Console.WriteLine(line);
}

if (wantSummary)
    await GenerateSummaryAsync(report, identical, different, missing);

return (different + missing) > 0 ? 1 : 0;

// -----------------------------------------------------------------------

static List<string> CompareExcelFiles(string path1, string path2)
{
    var cells1 = ReadAllCells(path1);
    var cells2 = ReadAllCells(path2);
    var diffs = new List<string>();

    foreach (var (sheet, cells) in cells1)
    {
        if (!cells2.TryGetValue(sheet, out var other))
        {
            diffs.Add($"{sheet}: sheet missing in second file");
            continue;
        }

        foreach (var (cell, val) in cells)
        {
            other.TryGetValue(cell, out string? otherVal);
            if (val != otherVal)
                diffs.Add($"{sheet} {cell}: [{val}] vs [{otherVal ?? ""}]");
        }

        foreach (var (cell, val) in other)
            if (!cells.ContainsKey(cell))
                diffs.Add($"{sheet} {cell}: [] vs [{val}]");
    }

    foreach (var sheet in cells2.Keys)
        if (!cells1.ContainsKey(sheet))
            diffs.Add($"{sheet}: sheet missing in first file");

    return diffs;
}

// Returns: sheet entry name → (cell reference → resolved text value)
static Dictionary<string, Dictionary<string, string>> ReadAllCells(string xlsxPath)
{
    var result = new Dictionary<string, Dictionary<string, string>>();

    using var zip = ZipFile.OpenRead(xlsxPath);

    // 1. Build shared strings table
    var sharedStrings = new List<string>();
    var ssEntry = zip.GetEntry("xl/sharedStrings.xml");
    if (ssEntry != null)
    {
        var doc = LoadXml(ssEntry);
        var ns = Ns(doc);
        foreach (XmlNode si in doc.SelectNodes("//x:si", ns)!)
        {
            var textNodes = si.SelectNodes("x:t | x:r/x:t", ns)!;
            sharedStrings.Add(string.Concat(textNodes.Cast<XmlNode>().Select(n => n.InnerText)));
        }
    }

    // 2. Read each worksheet
    var wsEntries = zip.Entries
        .Where(e => System.Text.RegularExpressions.Regex.IsMatch(
                        e.FullName, @"^xl/worksheets/sheet\d+\.xml$"))
        .OrderBy(e => e.FullName);

    foreach (var wsEntry in wsEntries)
    {
        var cells = new Dictionary<string, string>();
        var doc = LoadXml(wsEntry);
        var ns = Ns(doc);

        foreach (XmlNode c in doc.SelectNodes("//x:c", ns)!)
        {
            string? cellRef  = c.Attributes?["r"]?.Value;
            string? cellType = c.Attributes?["t"]?.Value;
            var vNode = c.SelectSingleNode("x:v", ns);
            if (cellRef == null || vNode == null) continue;

            string value = (cellType == "s" && int.TryParse(vNode.InnerText, out int idx) && idx < sharedStrings.Count)
                ? sharedStrings[idx]
                : vNode.InnerText;

            if (!string.IsNullOrEmpty(value))
                cells[cellRef] = value;
        }

        result[wsEntry.FullName] = cells;
    }

    return result;
}

static XmlDocument LoadXml(ZipArchiveEntry entry)
{
    using var stream = entry.Open();
    var doc = new XmlDocument();
    doc.Load(stream);
    return doc;
}

static XmlNamespaceManager Ns(XmlDocument doc)
{
    var ns = new XmlNamespaceManager(doc.NameTable);
    ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    return ns;
}

static async Task GenerateSummaryAsync(List<string> report, int identical, int different, int missing)
{
    Console.WriteLine("\n=== AI SUMMARY ===");

    string? apiKey = Environment.GetEnvironmentVariable("ANTHROPIC_API_KEY");
    if (string.IsNullOrEmpty(apiKey))
    {
        Console.WriteLine("[Skipped: ANTHROPIC_API_KEY environment variable is not set]");
        return;
    }

    // Limit prompt size — take up to 300 diff lines to stay well within token limits
    var diffLines = report.Take(300).ToList();
    bool truncated = report.Count > 300;

    string diffText = string.Join("\n", diffLines);
    if (truncated)
        diffText += $"\n... (truncated, {report.Count - 300} more lines not shown)";

    string prompt =
        $"The following are differences found when comparing two folders of Excel (.xlsx) files.\n" +
        $"Files: {identical} identical, {different} with differences, {missing} missing/extra.\n\n" +
        $"Differences:\n{diffText}\n\n" +
        $"Write a brief plain-language summary. Identify any patterns (e.g. the same type of " +
        $"difference in every file). State clearly whether the files are functionally equivalent " +
        $"or if there are meaningful data differences.";

    var requestBody = JsonSerializer.Serialize(new
    {
        model = "claude-haiku-4-5",
        max_tokens = 1024,
        messages = new[] { new { role = "user", content = prompt } }
    });

    using var http = new HttpClient();
    http.DefaultRequestHeaders.Add("x-api-key", apiKey);
    http.DefaultRequestHeaders.Add("anthropic-version", "2023-06-01");

    Console.Write("Contacting Claude... ");
    try
    {
        var response = await http.PostAsync(
            "https://api.anthropic.com/v1/messages",
            new StringContent(requestBody, Encoding.UTF8, "application/json"));

        string json = await response.Content.ReadAsStringAsync();

        if (!response.IsSuccessStatusCode)
        {
            Console.WriteLine($"[API error {(int)response.StatusCode}: {json}]");
            return;
        }

        using var doc = JsonDocument.Parse(json);
        string? text = doc.RootElement
            .GetProperty("content")[0]
            .GetProperty("text")
            .GetString();

        Console.WriteLine("done.\n");
        Console.WriteLine(text);
    }
    catch (Exception ex)
    {
        Console.WriteLine($"[Failed: {ex.Message}]");
    }
}
