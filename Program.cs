using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using ClosedXML.Excel;
using CsvHelper;

class Program
{
    static int Main(string[] args)
    {
        if (args.Length < 1)
        {
            PrintHelp();
            return 1;
        }

        try
        {
            var cmd = args[0].ToLowerInvariant();
            switch (cmd)
            {
                case "xlsx2csv":
                    if (args.Length < 3) { Console.WriteLine("xlsx2csv requires input.xlsx and output.csv"); return 1; }
                    XlsxToCsv(args[1], args[2]);
                    break;
                case "csv2xlsx":
                    if (args.Length < 3) { Console.WriteLine("csv2xlsx requires input.csv and output.xlsx"); return 1; }
                    CsvToXlsx(args[1], args[2]);
                    break;
                // case "mdb2csv":
                //     if (args.Length < 3) { Console.WriteLine("mdb2csv requires input.mdb and output-folder"); return 1; }
                //     MdbToCsv(args[1], args[2]);
                //     break;
                case "csv2xml":
                    if (args.Length < 3) { Console.WriteLine("csv2xml requires input.csv and output.xml"); return 1; }
                    CsvToXml(args[1], args[2]);
                    break;
                case "xml2csv":
                    if (args.Length < 3) { Console.WriteLine("xml2csv requires input.xml and output.csv"); return 1; }
                    XmlToCsv(args[1], args[2]);
                    break;
                default:
                    PrintHelp();
                    return 1;
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("Error: " + ex.Message);
            return 2;
        }

        return 0;
    }

    static void PrintHelp()
    {
        Console.WriteLine("C# File Converters\n");
        Console.WriteLine("Commands:");
        Console.WriteLine("  xlsx2csv <input.xlsx> <output.csv>");
        Console.WriteLine("  csv2xlsx <input.csv> <output.xlsx>");
        Console.WriteLine("  mdb2csv <input.mdb> <output-folder>");
        Console.WriteLine("  csv2xml  <input.csv> <output.xml>");
        Console.WriteLine("  xml2csv  <input.xml> <output.csv>");
    }

    static void XlsxToCsv(string xlsxPath, string csvPath)
    {
        using var wb = new XLWorkbook(xlsxPath);
        var ws = wb.Worksheets.First();

        using var writer = new StreamWriter(csvPath);
        using var csv = new CsvWriter(writer, CultureInfo.InvariantCulture);

        var firstRowUsed = ws.FirstRowUsed();
        if (firstRowUsed == null) return;

        var headerRow = firstRowUsed.RowUsed();
        var headers = new List<string>();
        foreach (var cell in headerRow.CellsUsed()) headers.Add(cell.GetString());
        foreach (var h in headers) csv.WriteField(h);
        csv.NextRecord();

        var rows = ws.RowsUsed().Skip(1);
        foreach (var row in rows)
        {
            foreach (var cell in row.Cells(1, headers.Count))
            {
                csv.WriteField(cell.Value.ToString() ?? string.Empty);
            }
            csv.NextRecord();
        }

        Console.WriteLine($"Saved CSV to {csvPath}");
    }

    static void CsvToXlsx(string csvPath, string xlsxPath)
    {
        using var reader = new StreamReader(csvPath);
        using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
        var records = new List<string[]>();
        while (csv.Read())
        {
            var row = new List<string>();
            for (int i = 0; csv.TryGetField<string>(i, out var field); i++) row.Add(field);
            records.Add(row.ToArray());
        }

        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");

        for (int r = 0; r < records.Count; r++)
        {
            var row = records[r];
            for (int c = 0; c < row.Length; c++)
            {
                ws.Cell(r + 1, c + 1).Value = row[c];
            }
        }

        wb.SaveAs(xlsxPath);
        Console.WriteLine($"Saved XLSX to {xlsxPath}");
    }

    // static void MdbToCsv(string mdbPath, string outputFolder)
    // {
    //     if (!Directory.Exists(outputFolder)) Directory.CreateDirectory(outputFolder);

    //     var providers = new[] { "Microsoft.ACE.OLEDB.12.0", "Microsoft.ACE.OLEDB.16.0", "Microsoft.Jet.OLEDB.4.0" };
    //     Exception lastEx = null;

    //     foreach (var provider in providers)
    //     {
    //         var connStr = $"Provider={provider};Data Source={mdbPath};Persist Security Info=False;";
    //         try
    //         {
    //             using var conn = new OleDbConnection(connStr);
    //             conn.Open();

    //             var tables = conn.GetSchema("Tables");
    //             foreach (DataRow row in tables.Rows)
    //             {
    //                 var tableType = row[3]?.ToString();
    //                 if (!string.Equals(tableType, "TABLE", StringComparison.OrdinalIgnoreCase)) continue;

    //                 var tableName = row[2].ToString();
    //                 var outCsv = Path.Combine(outputFolder, SanitizeFileName(tableName) + ".csv");
    //                 using var cmd = new OleDbCommand($"SELECT * FROM [{tableName}]", conn);
    //                 using var rdr = cmd.ExecuteReader();
    //                 using var writer = new StreamWriter(outCsv);
    //                 using var csv = new CsvWriter(writer, CultureInfo.InvariantCulture);

    //                 var schema = rdr.GetSchemaTable();
    //                 var cols = new List<string>();
    //                 foreach (DataRow c in schema.Rows) cols.Add(c[0].ToString());
    //                 foreach (var h in cols) csv.WriteField(h);
    //                 csv.NextRecord();

    //                 while (rdr.Read())
    //                 {
    //                     for (int i = 0; i < cols.Count; i++) csv.WriteField(rdr.IsDBNull(i) ? string.Empty : rdr.GetValue(i).ToString());
    //                     csv.NextRecord();
    //                 }

    //                 Console.WriteLine($"Exported table {tableName} -> {outCsv}");
    //             }

    //             return;
    //         }
    //         catch (Exception ex)
    //         {
    //             lastEx = ex;
    //         }
    //     }

    //     throw new InvalidOperationException("Failed to open MDB file with any known provider. Last error: " + lastEx?.Message, lastEx);
    // }

    static void CsvToXml(string csvPath, string xmlPath)
    {
        using var reader = new StreamReader(csvPath);
        using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
        csv.Read();
        csv.ReadHeader();
        var headers = csv.HeaderRecord;

        var doc = new XDocument(new XElement("Rows"));
        while (csv.Read())
        {
            var rowEl = new XElement("Row");
            foreach (var h in headers)
            {
                var val = csv.GetField(h);
                rowEl.Add(new XElement(XmlConvert.EncodeName(h), val));
            }
            doc.Root.Add(rowEl);
        }

        doc.Save(xmlPath);
        Console.WriteLine($"Saved XML to {xmlPath}");
    }

    static void XmlToCsv(string xmlPath, string csvPath)
    {
        var doc = XDocument.Load(xmlPath);
        var rows = doc.Root.Elements().ToList();
        if (!rows.Any())
        {
            File.WriteAllText(csvPath, string.Empty);
            return;
        }

        var headers = rows.SelectMany(r => r.Elements().Select(e => e.Name.LocalName)).Distinct().ToList();

        using var writer = new StreamWriter(csvPath);
        using var csv = new CsvWriter(writer, CultureInfo.InvariantCulture);

        foreach (var h in headers) csv.WriteField(h);
        csv.NextRecord();

        foreach (var r in rows)
        {
            foreach (var h in headers)
            {
                var el = r.Element(h);
                csv.WriteField(el?.Value ?? string.Empty);
            }
            csv.NextRecord();
        }

        Console.WriteLine($"Saved CSV to {csvPath}");
    }

    static string SanitizeFileName(string name)
    {
        foreach (var c in Path.GetInvalidFileNameChars()) name = name.Replace(c, '_');
        return name;
    }
}
