# DocxViewer

A small Windows desktop viewer for `.docx` files that never locks the file
against other users or apps, and renders the document faithfully (no
reformatting).

## Non-locking guarantee

The file is opened with:

```csharp
new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete)
```

`FileAccess.Read` means this process never writes to the file, and
`FileShare.ReadWrite | FileShare.Delete` means it never denies other
processes (including Word itself) read, write, or delete access while the
viewer has it open. Word's own collaborative editing/AutoSave lock file
(`~$filename.docx`) is untouched, since we don't create it.

## Formatting fidelity

Parsing is done directly against the OOXML package (`document.xml`,
`styles.xml`) via the OpenXML SDK, not by round-tripping through Word:

- **Paragraph marks (CR)**: each `w:p` becomes its own `Paragraph`, so
  paragraph breaks are preserved exactly as authored.
- **Line breaks (LF) / `w:br`, `w:cr`**: rendered as an explicit
  `LineBreak` inside the same paragraph, distinct from a paragraph mark.
- **Manual page breaks**: `w:pageBreakBefore` on a paragraph is mapped to
  `Paragraph.BreakPageBefore`, so pagination in `FlowDocumentPageViewer`
  matches Word's page breaks for that (most common) case. Other break
  types render as a visible line break.
- **Fonts, size, bold/italic/underline/strikethrough, color**: resolved
  per run with correct inheritance — run properties override the
  paragraph's style (walking the style's `w:basedOn` chain), which
  overrides the document's `w:docDefaults`.
- **Page size**: taken from the section's `w:pgSz` (twips) and applied to
  the `FlowDocument`'s page dimensions.
- **Tables**: basic structure (rows/cells/paragraphs) is preserved.
- **Tabs**: rendered as a tab character; whitespace in text runs is
  preserved as-is (`xml:space="preserve"` is honored because text is read
  verbatim, never trimmed).

Known simplification: a `w:br` of type `page`/`column` occurring
*mid-paragraph* (rather than via `w:pageBreakBefore`) is rendered as a
line break rather than a true page split, since WPF's `Paragraph` has no
inline hard-page-break primitive. Full-paragraph page breaks (the common
case for a deliberate page break) are exact.

## Build & run (Windows, .NET 8 SDK)

```
cd DocxViewer
dotnet run
```

Click **Open .docx...** and pick a file. It can be opened by other people
or by Word at the same time — this viewer only ever takes a shared read
handle.
