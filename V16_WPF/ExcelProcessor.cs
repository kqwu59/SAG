using System.Data;
using System.Globalization;
using System.Text;
using ClosedXML.Excel;

namespace NettoieXLSX.V16;

public sealed class ExcelProcessor
{
    private static readonly Dictionary<string, string[]> Synonyms = new(StringComparer.OrdinalIgnoreCase)
    {
        ["N° commande"] =
        [
            "n commande", "no commande", "numero commande", "n de commande", "n commande", "n° commande",
            "num commande", "n cmd", "no cmd", "numero cmd", "cmd", "commande", "order", "order id", "bdc"
        ],
        ["Libellé"] = ["libelle", "désignation", "designation", "objet", "description", "intitule", "intitulé", "libellé"],
        ["Fournisseur"] = ["fournisseur", "vendor", "tiers", "fournisseu"],
        ["Montant HT"] = ["montant ht", "total ht", "ht", "montant hors taxes", "m ht", "mnt ht", "montantht"],
        ["Ind. Visa"] = ["ind visa", "indice visa", "indicateur visa", "visa", "visa ind", "visa (ind)", "ind? visa", "ind.? visa"],
        ["Statut"] = ["statut", "status", "etat", "état"],
        ["Nature de dépense"] =
        [
            "nature de depense", "nature de dépense", "nature depense", "nature dépense",
            "nature de la depense", "nature de la dépense", "type de depense", "type de dépense", "nature"
        ],
        ["Type de flux"] = ["type de flux", "flux", "nature de flux"],
        ["Auteur"] = ["auteur", "saisi par", "cree par", "créé par"],
        ["Date de règlement"] =
        [
            "date de reglement", "date reglement", "date de paiement", "date paiement", "reglement", "paiement"
        ],
        ["Commande"] = ["commande", "n commande", "no commande", "numero commande", "n° commande", "cmd", "bdc"],
        ["Statut (constatations)"] = ["statut", "etat", "état"],
        ["Date"] = ["date", "date workflow", "workflow", "date de workflow", "dt workflow", "maj", "mise a jour", "mise à jour"]
    };

    public static void Process(
        string commandesPath,
        string? constatationsPath,
        string? facturesPath,
        string? envoiBdcPath,
        string? workflowPath,
        string outputPath,
        Action<string> log)
    {
        log("Lecture des fichiers...");
        var commandes = ProcessCommandes(commandesPath);
        var constatations = string.IsNullOrWhiteSpace(constatationsPath) ? null : ProcessConstatations(constatationsPath);
        var factures = string.IsNullOrWhiteSpace(facturesPath) ? null : ProcessFactures(facturesPath);
        var envoiBdc = string.IsNullOrWhiteSpace(envoiBdcPath) ? null : ProcessEnvoiBdc(envoiBdcPath);
        var workflow = string.IsNullOrWhiteSpace(workflowPath) ? null : ProcessWorkflow(workflowPath);

        log("Construction de l'onglet Global...");
        var global = CreateGlobal(commandes, envoiBdc, factures, workflow, constatations);

        log("Écriture du fichier de sortie...");
        WriteWorkbook(outputPath, commandes, envoiBdc, constatations, factures, workflow, global);

        log("Traitement terminé.");
    }

    private static DataTable ProcessCommandes(string path)
    {
        var data = ReadAfterSkip(path, 20);
        var fournisseurCol = PickColumn(data, "Fournisseur");
        var natureCol = PickColumn(data, "Nature de dépense");

        var rows = data.AsEnumerable();
        if (fournisseurCol is not null)
        {
            rows = rows.Where(row => !string.Equals(
                CleanUpper(row[fournisseurCol]),
                "FCM 3MUNDI ESR-M",
                StringComparison.OrdinalIgnoreCase));
        }

        if (natureCol is not null)
        {
            rows = rows.Where(row => NormalizeText(row[natureCol]) != "mission");
        }

        var output = CreateTable(
            "N° commande", "Libellé", "Fournisseur", "Montant HT", "Type de flux",
            "Nature de dépense", "Statut", "Ind. Visa", "Auteur");

        var columns = output.Columns;
        foreach (var row in rows)
        {
            var newRow = output.NewRow();
            newRow["N° commande"] = GetCellValue(row, PickColumn(data, "N° commande"));
            newRow["Libellé"] = GetCellValue(row, PickColumn(data, "Libellé"));
            newRow["Fournisseur"] = GetCellValue(row, PickColumn(data, "Fournisseur"));
            newRow["Montant HT"] = GetCellValue(row, PickColumn(data, "Montant HT"));
            newRow["Type de flux"] = GetCellValue(row, PickColumn(data, "Type de flux"));
            newRow["Nature de dépense"] = GetCellValue(row, PickColumn(data, "Nature de dépense"));
            newRow["Statut"] = GetCellValue(row, PickColumn(data, "Statut"));
            newRow["Ind. Visa"] = GetCellValue(row, PickColumn(data, "Ind. Visa"));
            newRow["Auteur"] = GetCellValue(row, PickColumn(data, "Auteur"));
            if (!RowIsEmpty(newRow, columns)) output.Rows.Add(newRow);
        }

        return output;
    }

    private static DataTable ProcessConstatations(string path)
    {
        var data = ReadAfterSkip(path, 17);
        var cmdCol = PickColumn(data, "Commande");
        var statutCol = PickColumn(data, "Statut (constatations)");

        var output = CreateTable("Commande", "extrait commande", "Statut");
        foreach (var row in data.AsEnumerable())
        {
            var commande = GetCellValue(row, cmdCol);
            var commandeText = CleanValue(commande);
            var newRow = output.NewRow();
            newRow["Commande"] = commande;
            newRow["extrait commande"] = string.IsNullOrWhiteSpace(commandeText) ? null : commandeText[..Math.Min(5, commandeText.Length)];
            newRow["Statut"] = GetCellValue(row, statutCol);
            if (!RowIsEmpty(newRow, output.Columns)) output.Rows.Add(newRow);
        }

        return output;
    }

    private static DataTable ProcessEnvoiBdc(string path)
    {
        var data = ReadAfterSkip(path, 0);
        var output = CreateTable("Commande", "Date envoi", "Agent");

        foreach (var row in data.AsEnumerable())
        {
            var newRow = output.NewRow();
            newRow["Commande"] = data.Columns.Count > 0 ? row[0] : null;
            newRow["Date envoi"] = data.Columns.Count > 1 ? row[1] : null;
            newRow["Agent"] = data.Columns.Count > 2 ? row[2] : null;
            if (!RowIsEmpty(newRow, output.Columns)) output.Rows.Add(newRow);
        }

        return output;
    }

    private static DataTable ProcessFactures(string path)
    {
        var data = ReadAfterSkip(path, 19);
        var natCol = PickColumn(data, "Nature de dépense");
        var fouCol = PickColumn(data, "Fournisseur");

        var rows = data.AsEnumerable();
        if (natCol is not null)
        {
            rows = rows.Where(row => !string.Equals(CleanUpper(row[natCol]), "MI", StringComparison.OrdinalIgnoreCase));
        }

        if (fouCol is not null)
        {
            rows = rows.Where(row => !string.Equals(
                CleanUpper(row[fouCol]),
                "FCM 3MUNDI ESR-M",
                StringComparison.OrdinalIgnoreCase));
        }

        var output = CreateTable("N° commande", "Montant HT", "Date de règlement");
        foreach (var row in rows)
        {
            var newRow = output.NewRow();
            newRow["N° commande"] = GetCellValue(row, PickColumn(data, "N° commande"));
            newRow["Montant HT"] = GetCellValue(row, PickColumn(data, "Montant HT"));
            newRow["Date de règlement"] = GetCellValue(row, PickColumn(data, "Date de règlement"));
            if (!RowIsEmpty(newRow, output.Columns)) output.Rows.Add(newRow);
        }

        return output;
    }

    private static DataTable ProcessWorkflow(string path)
    {
        return ReadBelowMarkerOrFirst(path, "Liste des résultats");
    }

    private static DataTable CreateGlobal(
        DataTable commandes,
        DataTable? envoiBdc,
        DataTable? factures,
        DataTable? workflow,
        DataTable? constatations)
    {
        var output = CreateTable("BDC", "OBJET", "FOURN.", "HT", "VISA", "ENVOYE", "SF", "WORKFLOW", "PAYE", "SOLDE", "STATUT");
        if (commandes.Rows.Count == 0) return output;

        var envoiLookup = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (envoiBdc is not null)
        {
            foreach (var row in envoiBdc.AsEnumerable())
            {
                var key = CleanValue(row["Commande"]);
                if (string.IsNullOrWhiteSpace(key) || envoiLookup.ContainsKey(key)) continue;
                var dateText = FormatDate(row["Date envoi"]);
                var agent = CleanValue(row["Agent"]);
                var combined = string.Join(" ", new[] { dateText, agent }.Where(s => !string.IsNullOrWhiteSpace(s)));
                envoiLookup[key] = combined.Trim();
            }
        }

        var factLookup = new Dictionary<string, FactureAggregate>(StringComparer.OrdinalIgnoreCase);
        if (factures is not null)
        {
            foreach (var row in factures.AsEnumerable())
            {
                var key = CleanValue(row["N° commande"]);
                if (string.IsNullOrWhiteSpace(key)) continue;
                if (!factLookup.TryGetValue(key, out var agg))
                {
                    agg = new FactureAggregate();
                    factLookup[key] = agg;
                }

                agg.Count += 1;
                agg.Sum += ToDecimal(row["Montant HT"]);
                if (agg.Count == 1)
                {
                    agg.Date = TryParseDate(row["Date de règlement"]);
                    agg.RawDate = CleanValue(row["Date de règlement"]);
                }
                else
                {
                    agg.Date = null;
                    agg.RawDate = null;
                }
            }
        }

        var workflowLookup = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (workflow is not null)
        {
            var (bdcCol, valueCol) = ChooseWorkflowColumns(workflow);
            if (bdcCol is not null && valueCol is not null)
            {
                foreach (var row in workflow.AsEnumerable())
                {
                    var key = CleanValue(row[bdcCol]);
                    if (string.IsNullOrWhiteSpace(key) || workflowLookup.ContainsKey(key)) continue;
                    var value = FormatDate(row[valueCol]) ?? CleanValue(row[valueCol]);
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        workflowLookup[key] = value;
                    }
                }
            }
        }

        var constatationByCmd = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var constatationByShort = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (constatations is not null)
        {
            foreach (var row in constatations.AsEnumerable())
            {
                var cmd = CleanValue(row["Commande"]);
                var extrait = CleanValue(row["extrait commande"]);
                var statut = CleanValue(row["Statut"]);
                if (!string.IsNullOrWhiteSpace(cmd) && !constatationByCmd.ContainsKey(cmd))
                {
                    constatationByCmd[cmd] = statut;
                }
                if (!string.IsNullOrWhiteSpace(extrait) && !constatationByShort.ContainsKey(extrait))
                {
                    constatationByShort[extrait] = statut;
                }
            }
        }

        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var row in commandes.AsEnumerable())
        {
            var bdc = CleanValue(row["N° commande"]);
            var objet = CleanValue(row["Libellé"]);
            var fournisseur = CleanValue(row["Fournisseur"]);
            var htValue = row["Montant HT"];
            var visa = CleanValue(row["Ind. Visa"]);
            var statut = CleanValue(row["Statut"]);

            var envoi = envoiLookup.TryGetValue(bdc, out var envoiValue) ? envoiValue : string.Empty;

            var sfValue = "Pas de SF connu";
            if (string.Equals(fournisseur, "BNP PARIBAS - REGULARISATION CARTE ACHAT", StringComparison.OrdinalIgnoreCase))
            {
                sfValue = "ss objet Régul CA";
            }
            else if (!string.IsNullOrWhiteSpace(envoi) &&
                     envoi.Contains("ss objet Régul CA", StringComparison.OrdinalIgnoreCase))
            {
                if (constatationByCmd.TryGetValue(bdc, out var stat))
                {
                    sfValue = stat;
                }
                else if (!string.IsNullOrWhiteSpace(bdc) &&
                         constatationByShort.TryGetValue(bdc[..Math.Min(5, bdc.Length)], out var statShort))
                {
                    sfValue = statShort;
                }
            }

            var workflowValue = workflowLookup.TryGetValue(bdc, out var wf) ? wf : string.Empty;

            var payeValue = "pas de paiement connu";
            var soldeValue = string.Empty;
            if (factLookup.TryGetValue(bdc, out var factAgg))
            {
                if (factAgg.Count == 1)
                {
                    if (factAgg.Date is not null)
                    {
                        payeValue = factAgg.Date.Value.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                    }
                    else if (!string.IsNullOrWhiteSpace(factAgg.RawDate))
                    {
                        payeValue = factAgg.RawDate;
                    }
                    else
                    {
                        payeValue = "date manquante";
                    }
                }
                else if (factAgg.Count > 1)
                {
                    payeValue = $"{factAgg.Count} paiement{(factAgg.Count > 1 ? "s" : string.Empty)}";
                }

                var solde = ToDecimal(htValue) - factAgg.Sum;
                soldeValue = solde.ToString("0.00", CultureInfo.InvariantCulture);
            }
            else if (!string.IsNullOrWhiteSpace(CleanValue(htValue)))
            {
                soldeValue = ToDecimal(htValue).ToString("0.00", CultureInfo.InvariantCulture);
            }

            var newRow = output.NewRow();
            newRow["BDC"] = bdc;
            newRow["OBJET"] = string.IsNullOrWhiteSpace(objet) ? "-" : objet;
            newRow["FOURN."] = string.IsNullOrWhiteSpace(fournisseur) ? "-" : fournisseur;
            newRow["HT"] = ToDecimal(htValue);
            newRow["VISA"] = string.IsNullOrWhiteSpace(visa) ? "-" : visa;
            newRow["ENVOYE"] = envoi;
            newRow["SF"] = sfValue;
            newRow["WORKFLOW"] = workflowValue;
            newRow["PAYE"] = payeValue;
            newRow["SOLDE"] = soldeValue;
            newRow["STATUT"] = string.IsNullOrWhiteSpace(statut) ? "-" : statut;

            var signature = string.Join("|", output.Columns.Cast<DataColumn>().Select(col => SigValue(newRow[col])));
            if (seen.Add(signature))
            {
                output.Rows.Add(newRow);
            }
        }

        return output;
    }

    private static (string? BdcColumn, string? ValueColumn) ChooseWorkflowColumns(DataTable table)
    {
        var bdc = PickColumn(table, "N° commande");
        var date = PickColumn(table, "Date");
        if (bdc is not null && date is not null) return (bdc, date);

        var statut = PickColumn(table, "Statut");
        if (bdc is not null && statut is not null) return (bdc, statut);

        if (bdc is not null && table.Columns.Count > 1) return (bdc, table.Columns[1].ColumnName);
        return (bdc, null);
    }

    private static void WriteWorkbook(
        string outputPath,
        DataTable commandes,
        DataTable? envoiBdc,
        DataTable? constatations,
        DataTable? factures,
        DataTable? workflow,
        DataTable global)
    {
        using var workbook = new XLWorkbook();
        AddSheet(workbook, "Commande", commandes);
        if (envoiBdc is not null) AddSheet(workbook, "Envoi BDC", envoiBdc);
        if (constatations is not null) AddSheet(workbook, "Constatation", constatations);
        if (factures is not null) AddSheet(workbook, "Factures", factures);
        if (workflow is not null) AddSheet(workbook, "Workflow", workflow);
        AddSheet(workbook, "Global", global);
        workbook.SaveAs(outputPath);
    }

    private static void AddSheet(XLWorkbook workbook, string name, DataTable table)
    {
        var sheet = workbook.Worksheets.Add(name);
        for (var col = 0; col < table.Columns.Count; col++)
        {
            sheet.Cell(1, col + 1).Value = table.Columns[col].ColumnName;
        }

        for (var row = 0; row < table.Rows.Count; row++)
        {
            for (var col = 0; col < table.Columns.Count; col++)
            {
                var value = table.Rows[row][col];
                if (value is DateTime dateValue)
                {
                    sheet.Cell(row + 2, col + 1).Value = dateValue.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                }
                else if (value is decimal decValue)
                {
                    sheet.Cell(row + 2, col + 1).Value = decValue;
                }
                else
                {
                    sheet.Cell(row + 2, col + 1).Value = value;
                }
            }
        }

        sheet.Columns().AdjustToContents();
    }

    private static DataTable ReadBelowMarkerOrFirst(string path, string marker)
    {
        using var workbook = new XLWorkbook(path);
        var markerNorm = NormalizeText(marker);
        foreach (var sheet in workbook.Worksheets)
        {
            foreach (var cell in sheet.CellsUsed())
            {
                if (NormalizeText(cell.GetString()) == markerNorm)
                {
                    return BuildTableFromWorksheet(sheet, cell.Address.RowNumber);
                }
            }
        }

        return ReadAfterSkip(path, 0);
    }

    private static DataTable ReadAfterSkip(string path, int skipRows)
    {
        using var workbook = new XLWorkbook(path);
        var sheet = workbook.Worksheets.First();
        return BuildTableFromWorksheet(sheet, skipRows + 1);
    }

    private static DataTable BuildTableFromWorksheet(IXLWorksheet sheet, int startRow)
    {
        var lastRow = sheet.LastRowUsed()?.RowNumber() ?? startRow;
        var lastCol = sheet.LastColumnUsed()?.ColumnNumber() ?? 1;

        var headerRow = -1;
        for (var row = startRow; row <= lastRow; row++)
        {
            var nonEmpty = 0;
            for (var col = 1; col <= lastCol; col++)
            {
                if (!string.IsNullOrWhiteSpace(sheet.Cell(row, col).GetString()))
                {
                    nonEmpty++;
                }
            }

            if (nonEmpty >= 2)
            {
                headerRow = row;
                break;
            }
        }

        var table = new DataTable();
        if (headerRow == -1) return table;

        var headers = new List<string>();
        for (var col = 1; col <= lastCol; col++)
        {
            var header = sheet.Cell(headerRow, col).GetString().Trim();
            if (string.IsNullOrWhiteSpace(header))
            {
                header = $"Column{col}";
            }

            header = DeduplicateHeader(headers, header);
            headers.Add(header);
            table.Columns.Add(header);
        }

        for (var row = headerRow + 1; row <= lastRow; row++)
        {
            var dataRow = table.NewRow();
            var hasValue = false;
            for (var col = 1; col <= lastCol; col++)
            {
                var cell = sheet.Cell(row, col);
                var value = cell.Value;
                if (value is double numeric && cell.DataType == XLDataType.DateTime)
                {
                    value = DateTime.FromOADate(numeric);
                }
                dataRow[col - 1] = value;
                if (!string.IsNullOrWhiteSpace(cell.GetString()))
                {
                    hasValue = true;
                }
            }

            if (hasValue) table.Rows.Add(dataRow);
        }

        return table;
    }

    private static string DeduplicateHeader(ICollection<string> existing, string header)
    {
        if (!existing.Contains(header)) return header;
        var index = 2;
        var candidate = $"{header} {index}";
        while (existing.Contains(candidate))
        {
            index++;
            candidate = $"{header} {index}";
        }
        return candidate;
    }

    private static DataTable CreateTable(params string[] columns)
    {
        var table = new DataTable();
        foreach (var column in columns)
        {
            table.Columns.Add(column);
        }
        return table;
    }

    private static string? PickColumn(DataTable table, string key)
    {
        if (!Synonyms.TryGetValue(key, out var synonyms)) return null;
        var normalized = table.Columns.Cast<DataColumn>()
            .ToDictionary(col => NormalizeText(col.ColumnName), col => col.ColumnName, StringComparer.OrdinalIgnoreCase);

        foreach (var synonym in synonyms)
        {
            var norm = NormalizeText(synonym);
            if (normalized.TryGetValue(norm, out var exact)) return exact;
        }

        foreach (var synonym in synonyms)
        {
            var norm = NormalizeText(synonym);
            foreach (var entry in normalized)
            {
                if (entry.Key.Contains(norm, StringComparison.OrdinalIgnoreCase))
                {
                    return entry.Value;
                }
            }
        }

        return null;
    }

    private static object? GetCellValue(DataRow row, string? column)
    {
        return column is null || !row.Table.Columns.Contains(column) ? null : row[column];
    }

    private static bool RowIsEmpty(DataRow row, DataColumnCollection columns)
    {
        foreach (DataColumn column in columns)
        {
            if (!string.IsNullOrWhiteSpace(CleanValue(row[column])))
            {
                return false;
            }
        }
        return true;
    }

    private static string CleanValue(object? value)
    {
        return value?.ToString()?.Trim() ?? string.Empty;
    }

    private static string CleanUpper(object? value)
    {
        return CleanValue(value).ToUpperInvariant();
    }

    private static string NormalizeText(object? value)
    {
        var text = CleanValue(value).ToLowerInvariant();
        var normalized = text.Normalize(NormalizationForm.FormD);
        var builder = new StringBuilder();
        foreach (var ch in normalized)
        {
            var category = CharUnicodeInfo.GetUnicodeCategory(ch);
            if (category == UnicodeCategory.NonSpacingMark) continue;
            if (char.IsLetterOrDigit(ch) || char.IsWhiteSpace(ch))
            {
                builder.Append(ch);
            }
        }
        return string.Join(' ', builder.ToString().Split(' ', StringSplitOptions.RemoveEmptyEntries));
    }

    private static decimal ToDecimal(object? value)
    {
        if (value is null) return 0m;
        if (value is decimal dec) return dec;
        if (value is double dbl) return (decimal)dbl;
        if (value is int i) return i;

        var text = CleanValue(value);
        text = text.Replace("\u00a0", string.Empty)
            .Replace("\u202f", string.Empty)
            .Replace("€", string.Empty)
            .Replace(" ", string.Empty)
            .Replace(',', '.');

        return decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out var result) ? result : 0m;
    }

    private static DateTime? TryParseDate(object? value)
    {
        if (value is DateTime dateTime) return dateTime.Date;
        if (value is double dbl) return DateTime.FromOADate(dbl).Date;

        var text = CleanValue(value);
        if (string.IsNullOrWhiteSpace(text)) return null;
        if (DateTime.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.None, out var parsed))
        {
            return parsed.Date;
        }

        if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out parsed))
        {
            return parsed.Date;
        }

        return null;
    }

    private static string? FormatDate(object? value)
    {
        var date = TryParseDate(value);
        return date?.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
    }

    private static string SigValue(object? value)
    {
        if (value is DateTime dt) return dt.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
        if (value is decimal dec) return dec.ToString("0.########", CultureInfo.InvariantCulture);
        return CleanValue(value);
    }

    private sealed class FactureAggregate
    {
        public int Count { get; set; }
        public decimal Sum { get; set; }
        public DateTime? Date { get; set; }
        public string? RawDate { get; set; }
    }
}
