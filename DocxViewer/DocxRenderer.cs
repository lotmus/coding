using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace DocxViewer;

/// <summary>
/// Converts a .docx package into a WPF FlowDocument, preserving line breaks (w:br),
/// paragraph marks, explicit page/column breaks, and per-run font/size/style/color,
/// without ever needing write access to the source file.
/// </summary>
internal static class DocxRenderer
{
    public static FlowDocument Render(Stream docxStream)
    {
        using var wordDocument = WordprocessingDocument.Open(docxStream, isEditable: false);

        var flowDocument = new FlowDocument
        {
            PagePadding = new Thickness(48),
            ColumnWidth = double.PositiveInfinity,
        };

        var mainPart = wordDocument.MainDocumentPart
            ?? throw new InvalidOperationException("The document has no main document part.");

        ApplyPageSize(flowDocument, mainPart.Document.Body);

        var defaults = RunFormat.FromDocDefaults(mainPart.StyleDefinitionsPart);
        var styles = mainPart.StyleDefinitionsPart?.Styles;

        foreach (var element in mainPart.Document.Body!.Elements())
        {
            switch (element)
            {
                case W.Paragraph paragraph:
                    flowDocument.Blocks.Add(RenderParagraph(paragraph, defaults, styles));
                    break;

                case W.Table table:
                    flowDocument.Blocks.Add(RenderTable(table, defaults, styles));
                    break;
            }
        }

        if (flowDocument.Blocks.Count == 0)
        {
            flowDocument.Blocks.Add(new Paragraph(new Run("(empty document)")));
        }

        return flowDocument;
    }

    private static void ApplyPageSize(FlowDocument flowDocument, W.Body? body)
    {
        var sectPr = body?.Elements<W.SectionProperties>().FirstOrDefault();
        var pageSize = sectPr?.GetFirstChild<W.PageSize>();
        if (pageSize?.Width?.Value is uint w && pageSize.Height?.Value is uint h)
        {
            // Page size is stored in twips (1/20 pt); WPF units are 1/96 in => twips / 15.
            flowDocument.PageWidth = w / 15.0;
            flowDocument.PageHeight = h / 15.0;
        }
    }

    private static Paragraph RenderParagraph(W.Paragraph paragraph, RunFormat defaults, W.Styles? styles)
    {
        var pPr = paragraph.ParagraphProperties;
        var result = new Paragraph
        {
            Margin = new Thickness(0),
        };

        ApplySpacing(result, pPr);
        ApplyAlignment(result, pPr);

        if (pPr?.PageBreakBefore != null)
        {
            result.BreakPageBefore = true;
        }

        var paragraphStyleFormat = RunFormat.FromParagraphStyle(paragraph, styles, defaults);

        foreach (var run in paragraph.Elements())
        {
            switch (run)
            {
                case W.Run r:
                    AppendRun(result, r, paragraphStyleFormat);
                    break;

                case W.Hyperlink hyperlink:
                    foreach (var hr in hyperlink.Elements<W.Run>())
                    {
                        AppendRun(result, hr, paragraphStyleFormat);
                    }
                    break;
            }
        }

        if (result.Inlines.Count == 0)
        {
            // Preserve genuinely empty paragraphs (blank lines) instead of collapsing them.
            result.Inlines.Add(new Run(string.Empty));
        }

        return result;
    }

    private static void AppendRun(Paragraph target, W.Run run, RunFormat paragraphFormat)
    {
        var format = RunFormat.FromRunProperties(run.RunProperties, paragraphFormat);

        foreach (var child in run.Elements())
        {
            switch (child)
            {
                case W.Text t:
                    // xml:space="preserve" is respected automatically because we read t.Text directly
                    // rather than trimming it.
                    target.Inlines.Add(CreateInline(t.Text, format));
                    break;

                case W.TabChar:
                    target.Inlines.Add(CreateInline("\t", format));
                    break;

                case W.Break:
                    // WPF's Paragraph has no inline hard-page-break primitive; mid-paragraph
                    // page/column breaks (w:br type="page"/"column") and plain line breaks
                    // all render as a visible line break here. Paragraph-level page breaks
                    // (w:pageBreakBefore, the common case for a manual page break) are
                    // reproduced exactly via Paragraph.BreakPageBefore in RenderParagraph.
                    target.Inlines.Add(new LineBreak());
                    break;

                case W.CarriageReturn:
                    target.Inlines.Add(new LineBreak());
                    break;
            }
        }
    }

    private static Inline CreateInline(string text, RunFormat format)
    {
        var run = new Run(text);
        format.ApplyTo(run);
        return run;
    }

    private static void ApplySpacing(Paragraph target, W.ParagraphProperties? pPr)
    {
        var spacing = pPr?.SpacingBetweenLines;
        double before = 0, after = 0;
        if (spacing?.Before?.Value != null && double.TryParse(spacing.Before.Value, out var b))
        {
            before = b / 20.0; // twentieths of a point -> points (~= px at 96/72 scale applied below)
        }
        if (spacing?.After?.Value != null && double.TryParse(spacing.After.Value, out var a))
        {
            after = a / 20.0;
        }
        target.Margin = new Thickness(0, before * (96.0 / 72.0), 0, after * (96.0 / 72.0));
    }

    private static void ApplyAlignment(Paragraph target, W.ParagraphProperties? pPr)
    {
        var jc = pPr?.Justification?.Val?.Value;
        target.TextAlignment = jc switch
        {
            W.JustificationValues.Center => TextAlignment.Center,
            W.JustificationValues.Right => TextAlignment.Right,
            W.JustificationValues.Both => TextAlignment.Justify,
            _ => TextAlignment.Left,
        };
    }

    private static Table RenderTable(W.Table table, RunFormat defaults, W.Styles? styles)
    {
        var wpfTable = new Table { CellSpacing = 0 };
        var rows = table.Elements<W.TableRow>().ToList();
        int columnCount = rows.Count == 0 ? 0 : rows.Max(r => r.Elements<W.TableCell>().Count());

        var group = new TableRowGroup();
        wpfTable.RowGroups.Add(group);

        for (int i = 0; i < columnCount; i++)
        {
            wpfTable.Columns.Add(new TableColumn());
        }

        foreach (var row in rows)
        {
            var wpfRow = new TableRow();
            foreach (var cell in row.Elements<W.TableCell>())
            {
                var wpfCell = new TableCell { BorderBrush = Brushes.Gray, BorderThickness = new Thickness(0.5), Padding = new Thickness(4) };
                foreach (var paragraph in cell.Elements<W.Paragraph>())
                {
                    wpfCell.Blocks.Add(RenderParagraph(paragraph, defaults, styles));
                }
                wpfRow.Cells.Add(wpfCell);
            }
            group.Rows.Add(wpfRow);
        }

        return wpfTable;
    }
}

/// <summary>Resolved (inherited) run-level formatting: font family, size, weight, style, color, underline.</summary>
internal sealed class RunFormat
{
    public string FontFamilyName { get; init; } = "Calibri";
    public double FontSizePt { get; init; } = 11;
    public bool Bold { get; init; }
    public bool Italic { get; init; }
    public bool Underline { get; init; }
    public bool Strike { get; init; }
    public Color? Color { get; init; }

    public void ApplyTo(Run run)
    {
        run.FontFamily = new FontFamily(FontFamilyName);
        run.FontSize = FontSizePt * (96.0 / 72.0);
        run.FontWeight = Bold ? FontWeights.Bold : FontWeights.Normal;
        run.FontStyle = Italic ? FontStyles.Italic : FontStyles.Normal;

        if (Underline || Strike)
        {
            var decorations = new TextDecorationCollection();
            if (Underline) decorations.Add(TextDecorations.Underline[0]);
            if (Strike) decorations.Add(TextDecorations.Strikethrough[0]);
            run.TextDecorations = decorations;
        }

        if (Color is Color c)
        {
            run.Foreground = new SolidColorBrush(c);
        }
    }

    public static RunFormat FromDocDefaults(StyleDefinitionsPart? stylesPart)
    {
        var rPrDefault = stylesPart?.Styles?.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle;
        var format = new RunFormat();
        return Merge(format, rPrDefault);
    }

    public static RunFormat FromParagraphStyle(W.Paragraph paragraph, W.Styles? styles, RunFormat inherited)
    {
        var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (styleId == null || styles == null)
        {
            return inherited;
        }

        var style = styles.Elements<W.Style>().FirstOrDefault(s => s.StyleId == styleId);
        var result = inherited;
        // Walk the style's BasedOn chain (base first) so more specific styles win.
        var chain = new List<W.Style>();
        var current = style;
        int guard = 0;
        while (current != null && guard++ < 20)
        {
            chain.Insert(0, current);
            var basedOnId = current.BasedOn?.Val?.Value;
            current = basedOnId == null ? null : styles.Elements<W.Style>().FirstOrDefault(s => s.StyleId == basedOnId);
        }
        foreach (var s in chain)
        {
            result = Merge(result, s.StyleRunProperties);
        }
        return result;
    }

    public static RunFormat FromRunProperties(W.RunProperties? rPr, RunFormat inherited) => Merge(inherited, rPr);

    private static RunFormat Merge(RunFormat baseFormat, DocumentFormat.OpenXml.OpenXmlCompositeElement? rPr)
    {
        if (rPr == null) return baseFormat;

        string fontFamily = baseFormat.FontFamilyName;
        var fonts = rPr.GetFirstChild<W.RunFonts>();
        if (fonts?.Ascii?.Value is string ascii) fontFamily = ascii;

        double fontSize = baseFormat.FontSizePt;
        var sz = rPr.GetFirstChild<W.FontSize>();
        if (sz?.Val?.Value is string szVal && double.TryParse(szVal, out var halfPoints))
        {
            fontSize = halfPoints / 2.0;
        }

        bool bold = baseFormat.Bold;
        var b = rPr.GetFirstChild<W.Bold>();
        if (b != null) bold = b.Val is null || b.Val.Value;

        bool italic = baseFormat.Italic;
        var i = rPr.GetFirstChild<W.Italic>();
        if (i != null) italic = i.Val is null || i.Val.Value;

        bool underline = baseFormat.Underline;
        var u = rPr.GetFirstChild<W.Underline>();
        if (u != null) underline = u.Val?.Value != null && u.Val.Value != W.UnderlineValues.None;

        bool strike = baseFormat.Strike;
        var s = rPr.GetFirstChild<W.Strike>();
        if (s != null) strike = s.Val is null || s.Val.Value;

        Color? color = baseFormat.Color;
        var colorEl = rPr.GetFirstChild<W.Color>();
        if (colorEl?.Val?.Value is string hex && hex != "auto" && TryParseHexColor(hex, out var parsed))
        {
            color = parsed;
        }

        return new RunFormat
        {
            FontFamilyName = fontFamily,
            FontSizePt = fontSize,
            Bold = bold,
            Italic = italic,
            Underline = underline,
            Strike = strike,
            Color = color,
        };
    }

    private static bool TryParseHexColor(string hex, out Color color)
    {
        color = default;
        if (hex.Length != 6) return false;
        try
        {
            byte r = Convert.ToByte(hex.Substring(0, 2), 16);
            byte g = Convert.ToByte(hex.Substring(2, 2), 16);
            byte bch = Convert.ToByte(hex.Substring(4, 2), 16);
            color = Color.FromRgb(r, g, bch);
            return true;
        }
        catch
        {
            return false;
        }
    }
}
