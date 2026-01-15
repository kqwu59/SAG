using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

namespace NettoieXLSX.V16
{
    public static class ExcelProcessor
    {
        private const string CoverText =
            "Document de référence — construction de l’onglet Global\n" +
            "Objectif : partir de vos fichiers sources, aligner les informations par numéro de commande (BDC) et remplir Global colonne par colonne.\n\n" +
            "1) D’où viennent les données (nettoyage de départ)\n" +
            "On ne garde que ce qu’il y a sous la ligne “Liste des résultats”.\n\n" +
            "• Feuille “Commande”\n" +
            "  Colonnes conservées : N° commande, Libellé, Fournisseur, Montant HT, Type de flux, Nature de dépense, Statut, Ind. Visa, Auteur.\n" +
            "  On retire les lignes où Fournisseur = FCM 3MUNDI ESR-M ou Nature de dépense = Mission.\n\n" +
            "• Feuille “Constatation”\n" +
            "  Colonnes : Commande, extrait commande (les 5 premiers caractères de Commande), Statut.\n\n" +
            "• Feuille “Envoi BDC”\n" +
            "  Colonnes : Commande, Date envoi, Agent.\n" +
            "  Les dates sont affichées sans l’heure.\n\n" +
            "• Feuille “Factures”\n" +
            "  Colonnes : N° commande, Montant HT, Date de règlement.\n" +
            "  On retire les lignes où Nature de dépense = MI ou Fournisseur = FCM 3MUNDI ESR-M.\n" +
            "  Les dates sont affichées sans l’heure.\n\n" +
            "• Feuille “Workflow”\n" +
            "  On garde tout. Plus tard, on y cherchera une colonne BDC (le numéro de commande) et une Date ou un Statut.\n\n" +
            "2) La clé qui relie tout : le BDC\n" +
            "Tout est relié avec le numéro de commande.\n\n" +
            "• Global.A (BDC) vient de Commande → N° commande.\n" +
            "• On cherche la même valeur :\n" +
            "  - dans Envoi BDC → Commande,\n" +
            "  - dans Factures → N° commande,\n" +
            "  - dans Workflow → (colonne BDC détectée automatiquement),\n" +
            "  - dans Constatation → Commande (ou, si rien, extrait commande = 5 premiers caractères du BDC).\n\n" +
            "Important : le BDC est traité comme du texte (Excel ne le transforme pas).\n\n" +
            "3) Comment on remplit chaque colonne de Global (A → K)\n" +
            "A — BDC : Commande → N° commande (texte)\n" +
            "B — OBJET : Commande → Libellé (\"-\" si vide)\n" +
            "C — FOURN. : Commande → Fournisseur (\"-\" si vide)\n" +
            "D — HT : Commande → Montant HT (base du Solde)\n" +
            "E — VISA : Commande → Ind. Visa (\"-\" si vide)\n" +
            "F — ENVOYE : Date envoi + Agent (Envoi BDC) si disponibles\n" +
            "G — SF : règle BNP Régul CA / Constatation sinon \"Pas de SF connu\"\n" +
            "H — WORKFLOW : Date (priorité) ou Statut depuis Workflow\n" +
            "I — PAYE : 0 facture → pas de paiement connu / 1 facture → date / ≥2 → n paiements\n" +
            "J — SOLDE : HT – somme des Montants HT factures\n" +
            "K — STATUT : Commande → Statut (\"-\" si vide)\n\n" +
            "4) Doublons\n" +
            "On supprime uniquement les lignes strictement identiques (A→K).";
        private const double CoverRowHeight = 18d;
        private const double CoverColumnWidth = 120d;
        private const double GlobalWidthOffset = 0.64d;
        private static readonly double[] GlobalColumnWidths =
        {
            7.09d, 36.09d, 70d, 12.09d, 16d, 14d, 16.82d, 30d, 8.09d, 8.09d, 12d
        };

        private sealed class Table
        {
            public List<string> Columns { get; } = new();
            public List<Dictionary<string, object?>> Rows { get; } = new();
        }

        private static readonly Dictionary<string, string[]> Synonyms = new()
        {
            ["N° commande"] = new[]
            {
                "n commande","no commande","numero commande","n de commande","n commande","n° commande",
                "num commande","n cmd","no cmd","numero cmd","cmd","commande","order","order id","bdc"
            },
            ["Libellé"] = new[] { "libelle","désignation","designation","objet","description","intitule","intitulé","libellé","objet" },
            ["Fournisseur"] = new[] { "fournisseur","vendor","tiers","fournisseu" },
            ["Montant HT"] = new[] { "montant ht","total ht","ht","montant hors taxes","m ht","mnt ht","montantht" },
            ["Ind. Visa"] = new[] { "ind visa","indice visa","indicateur visa","visa","visa ind","visa (ind)","ind? visa","ind.? visa" },
            ["Statut"] = new[] { "statut","status","etat","état" },
            ["Nature de dépense"] = new[]
            {
                "nature de depense","nature de dépense","nature depense","nature dépense",
                "nature de la depense","nature de la dépense","type de depense","type de dépense","nature"
            },
            ["Type de flux"] = new[] { "type de flux","flux","nature de flux" },
            ["Auteur"] = new[] { "auteur","saisi par","cree par","créé par" },
            ["Date de règlement"] = new[] { "date de reglement","date reglement","date de paiement","date paiement","reglement","paiement" },
            ["Commande"] = new[] { "commande","n commande","no commande","numero commande","n° commande","cmd","bdc" },
            ["Statut (constatations)"] = new[] { "statut","etat","état" },
            ["Date"] = new[] { "date","date workflow","workflow","date de workflow","dt workflow","maj","mise a jour","mise à jour" }
        };

        public static void Process(string commandesPath,
            string constatationsPath,
            string facturesPath,
            string envoiBdcPath,
            string workflowPath,
            string outputPath,
            Action<string>? log)
        {
            log?.Invoke("Lecture des sources...");

            var commandes = string.IsNullOrWhiteSpace(commandesPath) ? null : ProcessCommandes(commandesPath);
            var constatations = string.IsNullOrWhiteSpace(constatationsPath) ? null : ProcessConstatations(constatationsPath);
            var factures = string.IsNullOrWhiteSpace(facturesPath) ? null : ProcessFactures(facturesPath);
            var envoiBdc = string.IsNullOrWhiteSpace(envoiBdcPath) ? null : ProcessEnvoiBdc(envoiBdcPath);
            var workflow = string.IsNullOrWhiteSpace(workflowPath) ? null : ProcessWorkflow(workflowPath);

            log?.Invoke("Création du fichier de sortie...");
            using var output = new XLWorkbook();

            CreateCoverSheet(output);
            WriteTable(output, "Commande", commandes);
            WriteTable(output, "Envoi BDC", envoiBdc);
            WriteTable(output, "Constatation", constatations);
            WriteTable(output, "Factures", factures);
            WriteTable(output, "Workflow", workflow);
            CreateGlobalSheet(output, commandes, envoiBdc, factures, workflow, constatations);

            var outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrWhiteSpace(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            output.SaveAs(outputPath);
            log?.Invoke($"Fichier créé : {outputPath}");
        }

        private static Table ProcessCommandes(string path)
        {
            var table = ReadAfterSkip(path, 20);
            if (table.Columns.Count == 0)
            {
                return table;
            }

            var fournisseurCol = PickColumn(table.Columns, Synonyms["Fournisseur"]);
            var natureCol = PickColumn(table.Columns, Synonyms["Nature de dépense"]);
            var filtered = table.Rows.Where(row =>
            {
                var fournisseur = GetString(row, fournisseurCol);
                var nature = GetString(row, natureCol);
                if (!string.IsNullOrEmpty(fournisseur) &&
                    fournisseur.Trim().Equals("FCM 3MUNDI ESR-M", StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                if (!string.IsNullOrEmpty(nature) &&
                    Normalize(nature) == "mission")
                {
                    return false;
                }

                return true;
            }).ToList();

            var ordered = new[]
            {
                ("N° commande","N° commande"),
                ("Libellé","Libellé"),
                ("Fournisseur","Fournisseur"),
                ("Montant HT","Montant HT"),
                ("Type de flux","Type de flux"),
                ("Nature de dépense","Nature de dépense"),
                ("Statut","Statut"),
                ("Ind. Visa","Ind. Visa"),
                ("Auteur","Auteur"),
            };

            return ProjectTable(table, filtered, ordered);
        }

        private static Table ProcessConstatations(string path)
        {
            var table = ReadAfterSkip(path, 17);
            var cmdCol = PickColumn(table.Columns, Synonyms["Commande"]);
            var statutCol = PickColumn(table.Columns, Synonyms["Statut (constatations)"]);

            var output = new Table();
            output.Columns.AddRange(new[] { "Commande", "extrait commande", "Statut" });
            foreach (var row in table.Rows)
            {
                var cmd = GetValue(row, cmdCol);
                var cmdText = cmd?.ToString() ?? string.Empty;
                var shortCmd = cmdText.Length >= 5 ? cmdText[..5] : cmdText;
                output.Rows.Add(new Dictionary<string, object?>
                {
                    ["Commande"] = cmd,
                    ["extrait commande"] = shortCmd,
                    ["Statut"] = GetValue(row, statutCol)
                });
            }

            return output;
        }

        private static Table ProcessEnvoiBdc(string path)
        {
            var table = ReadAfterSkip(path, 0);
            if (table.Columns.Count == 0)
            {
                return table;
            }

            var firstThree = table.Columns.Take(3).ToList();
            while (firstThree.Count < 3)
            {
                firstThree.Add(string.Empty);
            }

            var output = new Table();
            output.Columns.AddRange(new[] { "Commande", "Date envoi", "Agent" });
            foreach (var row in table.Rows)
            {
                output.Rows.Add(new Dictionary<string, object?>
                {
                    ["Commande"] = GetValue(row, firstThree[0]),
                    ["Date envoi"] = GetValue(row, firstThree[1]),
                    ["Agent"] = GetValue(row, firstThree[2])
                });
            }

            return output;
        }

        private static Table ProcessFactures(string path)
        {
            var table = ReadAfterSkip(path, 19);
            var natureCol = PickColumn(table.Columns, Synonyms["Nature de dépense"]);
            var fournisseurCol = PickColumn(table.Columns, Synonyms["Fournisseur"]);
            var rows = table.Rows.Where(row =>
            {
                var nature = GetString(row, natureCol)?.Trim().ToUpperInvariant();
                if (nature == "MI")
                {
                    return false;
                }

                var fournisseur = GetString(row, fournisseurCol)?.Trim().ToUpperInvariant();
                if (fournisseur == "FCM 3MUNDI ESR-M")
                {
                    return false;
                }

                return true;
            }).ToList();

            var ordered = new[]
            {
                ("N° commande","N° commande"),
                ("Montant HT","Montant HT"),
                ("Date de règlement","Date de règlement")
            };

            return ProjectTable(table, rows, ordered);
        }

        private static Table ProcessWorkflow(string path)
        {
            var markerTable = ReadBelowMarker(path, "Liste des résultats");
            return markerTable.Columns.Count > 0 ? markerTable : ReadAfterSkip(path, 0);
        }

        private static Table ReadAfterSkip(string path, int skipRows)
        {
            using var workbook = new XLWorkbook(path);
            var ws = workbook.Worksheets.First();
            var rows = ws.RowsUsed()
                .Where(row => row.RowNumber() > skipRows)
                .ToList();
            var headerRow = rows.FirstOrDefault(row => row.CellsUsed().Count() >= 2);
            if (headerRow == null)
            {
                return new Table();
            }

            var lastColumn = ws.LastColumnUsed()?.ColumnNumber() ?? headerRow.LastCellUsed()?.Address.ColumnNumber ?? 0;
            if (lastColumn == 0)
            {
                return new Table();
            }

            var headers = new List<string>(lastColumn);
            for (var i = 1; i <= lastColumn; i++)
            {
                headers.Add(ws.Cell(headerRow.RowNumber(), i).GetValue<string>().Trim());
            }
            var headerIndex = headerRow.RowNumber();
            var table = new Table();
            table.Columns.AddRange(headers.Where(h => !string.IsNullOrWhiteSpace(h)));

            foreach (var row in ws.RowsUsed().Where(r => r.RowNumber() > headerIndex))
            {
                if (row.Cells().All(cell => cell.IsEmpty()))
                {
                    continue;
                }

                var dict = new Dictionary<string, object?>();
                for (var i = 0; i < headers.Count; i++)
                {
                    var header = headers[i];
                    if (string.IsNullOrWhiteSpace(header))
                    {
                        continue;
                    }

                    dict[header] = ws.Cell(row.RowNumber(), i + 1).Value;
                }

                if (dict.Count > 0)
                {
                    table.Rows.Add(dict);
                }
            }

            return table;
        }

        private static void CreateCoverSheet(XLWorkbook workbook)
        {
            var ws = workbook.AddWorksheet("Page de garde");
            var lines = CoverText.Split('\n');
            for (var i = 0; i < lines.Length; i++)
            {
                ws.Cell(i + 1, 1).Value = lines[i];
                ws.Cell(i + 1, 1).Style.Alignment.WrapText = true;
                ws.Cell(i + 1, 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                ws.Row(i + 1).Height = CoverRowHeight;
            }

            ws.Column(1).Width = CoverColumnWidth;
            ws.SheetView.FreezeRows(1);
        }

        private static Table ReadBelowMarker(string path, string marker)
        {
            using var workbook = new XLWorkbook(path);
            foreach (var ws in workbook.Worksheets)
            {
                foreach (var cell in ws.CellsUsed())
                {
                    var text = cell.GetValue<string>();
                    if (Normalize(text) == Normalize(marker))
                    {
                        var startRow = cell.Address.RowNumber;
                        var rows = ws.RowsUsed().Where(r => r.RowNumber() >= startRow).ToList();
                        var headerRow = rows.FirstOrDefault(r => r.CellsUsed().Count() >= 2);
                        if (headerRow == null)
                        {
                            return new Table();
                        }

                        var lastColumn = ws.LastColumnUsed()?.ColumnNumber() ?? headerRow.LastCellUsed()?.Address.ColumnNumber ?? 0;
                        if (lastColumn == 0)
                        {
                            return new Table();
                        }

                        var headers = new List<string>(lastColumn);
                        for (var i = 1; i <= lastColumn; i++)
                        {
                            headers.Add(ws.Cell(headerRow.RowNumber(), i).GetValue<string>().Trim());
                        }
                        var table = new Table();
                        table.Columns.AddRange(headers.Where(h => !string.IsNullOrWhiteSpace(h)));
                        var headerIndex = headerRow.RowNumber();
                        foreach (var row in ws.RowsUsed().Where(r => r.RowNumber() > headerIndex))
                        {
                            if (row.Cells().All(c => c.IsEmpty()))
                            {
                                continue;
                            }

                            var dict = new Dictionary<string, object?>();
                            for (var i = 0; i < headers.Count; i++)
                            {
                                var header = headers[i];
                                if (string.IsNullOrWhiteSpace(header))
                                {
                                    continue;
                                }

                                dict[header] = ws.Cell(row.RowNumber(), i + 1).Value;
                            }

                            if (dict.Count > 0)
                            {
                                table.Rows.Add(dict);
                            }
                        }

                        return table;
                    }
                }
            }

            return new Table();
        }

        private static void WriteTable(XLWorkbook workbook, string name, Table? table)
        {
            if (table == null || table.Columns.Count == 0)
            {
                return;
            }

            var ws = workbook.AddWorksheet(name);
            for (var i = 0; i < table.Columns.Count; i++)
            {
                ws.Cell(1, i + 1).Value = table.Columns[i];
            }

            for (var r = 0; r < table.Rows.Count; r++)
            {
                var row = table.Rows[r];
                for (var c = 0; c < table.Columns.Count; c++)
                {
                    row.TryGetValue(table.Columns[c], out var value);
                    var cell = ws.Cell(r + 2, c + 1);
                    cell.Value = XLCellValue.FromObject(value ?? string.Empty);
                    if (value is DateTime dateValue)
                    {
                        cell.Style.DateFormat.Format = "dd/MM/yyyy";
                    }
                }
            }

            AutoFitWorksheet(ws, table);
        }

        private static void AutoFitWorksheet(IXLWorksheet ws, Table table, int minWidth = 10, int maxWidth = 60, int padding = 2)
        {
            for (var i = 0; i < table.Columns.Count; i++)
            {
                var header = table.Columns[i] ?? string.Empty;
                var maxLen = header.Length;
                foreach (var row in table.Rows)
                {
                    if (row.TryGetValue(table.Columns[i], out var value) && value != null)
                    {
                        var length = value.ToString()?.Length ?? 0;
                        if (length > maxLen)
                        {
                            maxLen = length;
                        }
                    }
                }

                var width = Math.Min(maxWidth, Math.Max(minWidth, maxLen + padding));
                ws.Column(i + 1).Width = width;
            }
        }

        private static void CreateGlobalSheet(XLWorkbook workbook,
            Table? commandes,
            Table? envoiBdc,
            Table? factures,
            Table? workflow,
            Table? constatations)
        {
            var ws = workbook.AddWorksheet("Global");
            var headers = new[]
            {
                "BDC", "OBJET", "FOURN.", "HT", "VISA", "ENVOYE",
                "SF", "WORKFLOW", "PAYE", "SOLDE", "STATUT"
            };
            for (var i = 0; i < headers.Length; i++)
            {
                ws.Cell(1, i + 1).Value = headers[i];
            }
            for (var i = 0; i < GlobalColumnWidths.Length; i++)
            {
                ws.Column(i + 1).Width = GlobalColumnWidths[i] + GlobalWidthOffset;
            }

            var headerRow = ws.Row(1);
            headerRow.Height = 30d;
            headerRow.Style.Font.FontName = "Calibri";
            headerRow.Style.Font.FontSize = 12d;
            headerRow.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            headerRow.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            if (commandes == null || commandes.Columns.Count == 0)
            {
                return;
            }

            var envoiLookup = new Dictionary<string, string>();
            if (envoiBdc != null)
            {
                foreach (var row in envoiBdc.Rows)
                {
                    var key = GetString(row, "Commande")?.Trim();
                    if (string.IsNullOrWhiteSpace(key))
                    {
                        continue;
                    }

                    var dateText = FormatDate(GetValue(row, "Date envoi"));
                    var agent = GetString(row, "Agent")?.Trim() ?? string.Empty;
                    var value = $"{dateText} {agent}".Trim();
                    if (!envoiLookup.ContainsKey(key))
                    {
                        envoiLookup[key] = value;
                    }
                }
            }

            var factAgg = new Dictionary<string, (int Count, decimal Sum, object? RawDate, DateTime? ParsedDate)>();
            if (factures != null)
            {
                foreach (var row in factures.Rows)
                {
                    var key = GetString(row, "N° commande")?.Trim();
                    if (string.IsNullOrWhiteSpace(key))
                    {
                        continue;
                    }

                    var amount = ToDecimal(GetValue(row, "Montant HT"));
                    var rawDate = GetValue(row, "Date de règlement");
                    var parsed = ToDate(rawDate);

                    if (!factAgg.TryGetValue(key, out var existing))
                    {
                        existing = (0, 0m, null, null);
                    }

                    existing.Count += 1;
                    existing.Sum += amount;
                    if (existing.Count == 1)
                    {
                        existing.RawDate = rawDate;
                        existing.ParsedDate = parsed;
                    }
                    else
                    {
                        existing.RawDate = null;
                        existing.ParsedDate = null;
                    }

                    factAgg[key] = existing;
                }
            }

            var wfLookup = new Dictionary<string, object?>();
            if (workflow != null)
            {
                var bdcCol = PickColumn(workflow.Columns, Synonyms["N° commande"]);
                var valueCol = PickColumn(workflow.Columns, Synonyms["Date"]) ??
                               PickColumn(workflow.Columns, Synonyms["Statut"]);
                foreach (var row in workflow.Rows)
                {
                    var key = GetString(row, bdcCol)?.Trim();
                    if (string.IsNullOrWhiteSpace(key))
                    {
                        continue;
                    }

                    wfLookup[key] = valueCol != null ? GetValue(row, valueCol) : null;
                }
            }

            var constFull = new Dictionary<string, object?>();
            var constShort = new Dictionary<string, object?>();
            if (constatations != null)
            {
                foreach (var row in constatations.Rows)
                {
                    var full = GetString(row, "Commande")?.Trim();
                    var shortKey = GetString(row, "extrait commande")?.Trim();
                    var status = GetValue(row, "Statut");
                    if (!string.IsNullOrWhiteSpace(full))
                    {
                        constFull[full] = status;
                    }

                    if (!string.IsNullOrWhiteSpace(shortKey))
                    {
                        constShort[shortKey] = status;
                    }
                }
            }

            var seen = new HashSet<string>();
            var rowIndex = 2;
            foreach (var row in commandes.Rows)
            {
                var bdc = GetString(row, "N° commande")?.Trim();
                if (string.IsNullOrWhiteSpace(bdc))
                {
                    continue;
                }

                var objet = GetString(row, "Libellé") ?? "-";
                var fournisseur = GetString(row, "Fournisseur") ?? "-";
                var htValue = GetValue(row, "Montant HT");
                var visa = GetString(row, "Ind. Visa") ?? "-";
                var envoi = envoiLookup.TryGetValue(bdc, out var envoiValue) ? envoiValue : string.Empty;

                var sf = "Pas de SF connu";
                if (string.Equals(fournisseur?.Trim(), "BNP PARIBAS - REGULARISATION CARTE ACHAT",
                        StringComparison.OrdinalIgnoreCase))
                {
                    sf = "ss objet Régul CA";
                }
                else if (Normalize(envoi).Contains("ss objet regul ca"))
                {
                    if (constFull.TryGetValue(bdc, out var status) || constShort.TryGetValue(bdc[..Math.Min(5, bdc.Length)], out status))
                    {
                        sf = status?.ToString() ?? "Pas de SF connu";
                    }
                }

                wfLookup.TryGetValue(bdc, out var workflowValue);
                var workflowDisplay = workflowValue ?? string.Empty;

                var paye = "pas de paiement connu";
                if (factAgg.TryGetValue(bdc, out var agg))
                {
                    if (agg.Count == 1)
                    {
                        paye = agg.ParsedDate.HasValue ? agg.ParsedDate.Value.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture)
                            : agg.RawDate?.ToString() ?? "date manquante";
                    }
                    else
                    {
                        paye = $"{agg.Count} paiement{(agg.Count >= 2 ? "s" : string.Empty)}";
                    }
                }

                var solde = ToDecimal(htValue) - (factAgg.TryGetValue(bdc, out agg) ? agg.Sum : 0m);
                var statut = GetString(row, "Statut") ?? "-";

                var values = new object?[]
                {
                    bdc, objet, fournisseur, htValue ?? string.Empty, visa, envoi, sf,
                    workflowDisplay, paye, solde, statut
                };
                var signature = string.Join("|", values.Select(SignatureValue));
                if (!seen.Add(signature))
                {
                    continue;
                }

                for (var c = 0; c < values.Length; c++)
                {
                    var cell = ws.Cell(rowIndex, c + 1);
                    cell.Value = XLCellValue.FromObject(values[c] ?? string.Empty);
                    cell.Style.Font.FontName = "Calibri";
                    cell.Style.Font.FontSize = 9d;
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    if (values[c] is DateTime dateVal)
                    {
                        cell.Style.DateFormat.Format = "dd/MM/yyyy";
                    }
                }

                ws.Row(rowIndex).Height = 30d;
                ws.Cell(rowIndex, 1).Style.NumberFormat.Format = "@";
                ws.Cell(rowIndex, 10).Style.NumberFormat.Format = "0.00";
                rowIndex++;
            }
        }

        private static Table ProjectTable(Table source, List<Dictionary<string, object?>> rows,
            (string Target, string SynKey)[] order)
        {
            var output = new Table();
            foreach (var (target, synKey) in order)
            {
                output.Columns.Add(target);
                var column = PickColumn(source.Columns, Synonyms[synKey]);
                foreach (var row in rows)
                {
                    if (!row.ContainsKey(target))
                    {
                        row[target] = column != null ? GetValue(row, column) : null;
                    }
                }
            }

            foreach (var row in rows)
            {
                var projected = new Dictionary<string, object?>();
                foreach (var (target, _) in order)
                {
                    row.TryGetValue(target, out var value);
                    projected[target] = value;
                }
                output.Rows.Add(projected);
            }

            return output;
        }

        private static string? PickColumn(IEnumerable<string> columns, IEnumerable<string> synonyms)
        {
            var map = new Dictionary<string, string>();
            foreach (var col in columns)
            {
                var key = Normalize(col);
                if (string.IsNullOrWhiteSpace(key))
                {
                    continue;
                }

                if (!map.ContainsKey(key))
                {
                    map[key] = col;
                }
            }
            foreach (var syn in synonyms)
            {
                var normalized = Normalize(syn);
                if (map.TryGetValue(normalized, out var col))
                {
                    return col;
                }
            }

            foreach (var syn in synonyms)
            {
                var normalized = Normalize(syn);
                var match = map.Keys.FirstOrDefault(k => k.Contains(normalized));
                if (match != null)
                {
                    return map[match];
                }
            }

            return null;
        }

        private static object? GetValue(Dictionary<string, object?> row, string? column)
        {
            if (string.IsNullOrWhiteSpace(column))
            {
                return null;
            }

            row.TryGetValue(column, out var value);
            return value;
        }

        private static string? GetString(Dictionary<string, object?> row, string? column)
            => GetValue(row, column)?.ToString();

        private static string Normalize(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var normalized = value.Normalize(NormalizationForm.FormD);
            var builder = new StringBuilder();
            foreach (var ch in normalized)
            {
                var unicode = CharUnicodeInfo.GetUnicodeCategory(ch);
                if (unicode != UnicodeCategory.NonSpacingMark)
                {
                    builder.Append(ch);
                }
            }

            var cleaned = new string(builder.ToString().ToLowerInvariant().Select(ch => char.IsLetterOrDigit(ch) ? ch : ' ').ToArray());
            return string.Join(' ', cleaned.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries));
        }

        private static decimal ToDecimal(object? value)
        {
            if (value == null)
            {
                return 0m;
            }

            if (value is decimal dec)
            {
                return dec;
            }

            if (value is double dbl)
            {
                return Convert.ToDecimal(dbl);
            }

            if (value is int i)
            {
                return i;
            }

            var text = value.ToString() ?? string.Empty;
            text = text.Replace("\u00a0", string.Empty).Replace("\u202f", string.Empty);
            text = text.Replace("€", string.Empty).Replace(" ", string.Empty).Replace(",", ".");
            return decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out var result) ? result : 0m;
        }

        private static DateTime? ToDate(object? value)
        {
            if (value is DateTime dt)
            {
                return dt.Date;
            }

            if (value == null)
            {
                return null;
            }

            if (DateTime.TryParse(value.ToString(), new CultureInfo("fr-FR"), DateTimeStyles.None, out var parsed))
            {
                return parsed.Date;
            }

            return null;
        }

        private static string FormatDate(object? value)
        {
            var date = ToDate(value);
            return date.HasValue ? date.Value.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture) : value?.ToString() ?? string.Empty;
        }

        private static string SignatureValue(object? value)
        {
            if (value is DateTime date)
            {
                return date.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
            }

            if (value is decimal dec)
            {
                return dec.ToString(CultureInfo.InvariantCulture);
            }

            if (value is double dbl)
            {
                return dbl.ToString("G", CultureInfo.InvariantCulture);
            }

            return value?.ToString() ?? string.Empty;
        }
    }
}
