# Layout → DOCX (`from-layout`)

`from-layout` builds a brand new `.docx` from an **absolute** layout description, typically
produced by a Typst renderer (e.g. a CV rendered in Rust). The goal is a "pixel perfect" Word
export: every glyph line, rule, box and picture is placed where the layout engine put it,
instead of letting Word re-flow the text.

## Entry points

| Surface | Call |
|---------|------|
| CLI | `docx-cli from-layout <layout.json\|-> -o <out.docx> [--baseline-ratio R] [--slack PT]` |
| Library | `LayoutDocxWriter.Write(LayoutDocument, Stream[, LayoutDocxOptions])`, `LayoutParser.Parse(string)` (namespace `DocxMcp.Layout`) |
| MCP | `document_from_layout(layout_json, output_path?, baseline_ratio?)` — writes the file, or returns base64 when `output_path` is omitted |

The CLI command is handled **before** the session / storage / gRPC bootstrap, so it runs
standalone (no sessions directory, no storage server). `-` reads the layout from stdin.

```bash
docx-cli from-layout cv.layout.json -o cv.docx
cat cv.layout.json | docx-cli from-layout - -o cv.docx --baseline-ratio 0.8
```

## Schema (version 1)

Units are typographic points (pt, 1/72 in). Origin is the **top-left** corner of the page and
y grows downward. Colors are `"#RRGGBB"` (`#RGB` and `#RRGGBBAA` are accepted; alpha is dropped).

```jsonc
{
  "version": 1,
  "pages": [
    {
      "width": 595.28, "height": 841.89,
      "items": [                       // paint order: later items are on top
        { "type": "text",
          "x": 70.9, "y": 100.2, "width": 300.0, "height": 40.5,   // block box; y = top of first line box
          "lines": [
            { "baseline": 12.1,        // baseline offset from the block top
              "height": 14.0,          // distance to the next baseline (last line: font line height)
              "align": "left",         // left | right | center | justify
              "runs": [
                { "text": "Hello ", "font": "General Sans Medium", "size": 10.0,
                  "bold": false, "italic": false, "color": "#515151",
                  "underline": false, "link": null,                  // link: URL or null
                  "spacing": 2.35 } ] } ] },                         // optional, see below
        { "type": "rect", "x": 0, "y": 0, "width": 100, "height": 20,
          "fill": "#1C1C1C", "stroke": null },                     // stroke: {"color","width"} or null
        { "type": "line", "x1": 70, "y1": 800, "x2": 525, "y2": 800,
          "stroke": {"color": "#FCB912", "width": 0.6} },
        { "type": "image", "x": 420, "y": 30, "width": 102, "height": 55,
          "format": "png", "data": "<base64>" }                     // png | jpeg (gif, bmp, tiff too)
      ] } ] }
```

Unknown properties are ignored; unknown item types, unknown `align` values, invalid colors and
versions other than 1 are rejected with a message naming the offending path
(e.g. `pages[0].items[3].lines[1]`). A sample covering every item type lives in
`tests/DocxMcp.Tests/TestData/layout-sample.json`.

## Output structure

- **One section per layout page**, with the page's exact size (twips, rounded to 1/20 pt),
  all margins, header and footer distances 0, and a `nextPage` break between pages.
- **One tiny anchor paragraph per page** (1 pt font, exact 1 pt line, no spacing) — nothing can
  overflow onto an extra page. For every page but the last, the paragraph carries the section's
  `w:sectPr`; the last page uses the body `w:sectPr`.
- **Every item is an anchored DrawingML object** (`wp:anchor`, `behindDoc=0`, `allowOverlap=1`,
  `wrapNone`, `positionH`/`positionV` `relativeFrom="page"`, offsets in EMU: 1 pt = 12700 EMU),
  in paint order. `relativeHeight` follows Word's own scheme (251659264 + 1024·n): Word does not
  honour small values such as 1, 2, 3 and would paint the first object on top.
  - `text` → a text box (`wps:wsp` + `wps:txbx`/`w:txbxContent`, no fill, no line, insets 0,
    `noAutofit`). Each layout line is one paragraph: exact line spacing, `before`/`after` 0,
    alignment per line (`justify` → `w:jc="distribute"`, which stretches a single line to the
    full width), runs with `w:rFonts` (ascii/hAnsi/eastAsia/cs), `w:sz` (half-points), `w:b`,
    `w:i`, `w:color`, `w:u="single"`, kerning at every size and standard ligatures (as Typst);
    an optional run `spacing` (pt, may be negative: extra advance after each character) becomes
    `w:spacing w:val="round(spacing×20)"` in the run, clamped to Word's ±1584 pt — this lets the
    exporter reproduce Typst's justification exactly (one run per stretched space, line
    `align: "left"`) instead of relying on `distribute`, which Word spreads differently (partly
    between letters); links become `w:hyperlink` with an external relationship (one per distinct URL); spaces are
    preserved.
  - `rect` → `wps:wsp`, `prstGeom rect`, `solidFill`/`noFill`, `a:ln` (width/color or `noFill`).
  - `line` → `wps:wsp`, `prstGeom line`, box = bounding box of the two points, `flipH` when the
    line goes up to the right; flat caps.
  - `image` → `pic:pic` with the image part embedded, stretched to the exact box. Unsupported
    formats (SVG…) are skipped with a warning on stderr (an empty shape keeps the paint order).
- Fonts are referenced by name only (no embedding): install the fonts used by the layout.
- Plain `wp:anchor` + `wps` (no `mc:AlternateContent` / VML fallback): Word 2010+ and
  LibreOffice read it. `mc:Ignorable="w14 wp14"` is declared.
- Minimal `styles.xml` (doc defaults with no spacing) and `settings.xml`
  (`compatibilityMode` 15).
- **Deterministic**: fixed relationship ids, fixed docPr ids, and zip entries rewritten with a
  fixed 1980-01-01 timestamp: the same input gives the same bytes.

## Placement rules

### Horizontal

A text box is widened by a *slack* (default 24 pt, `--slack`) on each side, and each paragraph
is indented back so the text stays at the layout position:

| align | box extends | paragraph indents |
|-------|-------------|-------------------|
| left | both sides | left = slack, right = 0 → text starts at `x`, may overflow right by `slack` |
| right | both sides | left = 0, right = slack → text ends at `x + width` |
| center | both sides | 0 / 0 → centered on the block center |
| justify | both sides | slack / slack → stretched to exactly `width` (`distribute`) |

Prefer exporting justified lines as `align: "left"` with per-run `spacing` on the stretched
spaces: Word's `distribute` does not spread the space the way Typst does.

This keeps a line that Word measures marginally wider than Typst from wrapping.

### Vertical (baseline)

Word positions text inside an exact-height line itself; the writer chooses the text-box top and
each line's exact height so that every baseline lands on `y + lines[k].baseline`:

1. First line: exact height `h0 = max(lines[0].height, font size)`; box top =
   `y + lines[0].baseline − BaselineOffset(h0)`.
2. Line k > 0: its top is the previous line's bottom; its exact height is solved so that
   `top_k + BaselineOffset(h_k) = y + lines[k].baseline`. Heights are rounded to twips but each
   line is solved against its absolute target, so errors never accumulate.
3. The box is made taller than the content (by the last font size) because Word hides text that
   overflows a text box.

All the Word-specific knowledge is in **one function**, `BaselineModel.BaselineOffset(lineHeight,
fontSizePt)` (`src/DocxMcp/Layout/BaselineModel.cs`):

```
BaselineOffset(lineHeight, size) = ratio × lineHeight        ratio = 0.8
```

Calibration: single-line text boxes in Helvetica, Arial, Times New Roman and Georgia at
8/10/12/24 pt, exact line heights from 0.5× to 3× the font size, exported to PDF by Word for Mac
(compatibility mode 15) and measured from the PDF glyph origins: the baseline is at
0.80 × line height in every case, within the 0.24 pt resolution of Word's PDF output —
independent of the font and of the font size. LibreOffice places it identically.

### Calibration knob

| Knob | Effect |
|------|--------|
| `--baseline-ratio R` (CLI) | ratio used by `BaselineOffset` |
| `DOCX_LAYOUT_BASELINE_RATIO=R` (env, CLI and MCP) | same, when the flag is absent |
| `baseline_ratio` (MCP argument) | same, per call |
| `--slack PT` (CLI) / `LayoutDocxOptions.HorizontalSlack` | horizontal slack of text boxes |

To re-calibrate (another Word version, platform or compatibility mode): generate a layout with
single-line blocks of known `y`/`baseline`/`height`, convert the `.docx` to PDF, read the glyph
baselines from the PDF (e.g. PyMuPDF `page.get_text("rawdict")` span origins) and compare with
`y + baseline`; a uniform error `e` on lines of height `h` means `ratio += e / h`.

## Visual checks

Visual checks are done by converting the generated `.docx` to PDF and comparing it with the
Typst PDF, e.g. with LibreOffice:

```bash
soffice --headless --convert-to pdf --outdir out/ cv.docx
pdftoppm -r 100 -png out/cv.pdf out/page
```

## Known limitations

- Font sizes are rounded to half-points (`w:sz`), so a 9.8 pt run becomes 10 pt.
- Word and Typst may measure the same text slightly differently (hinting, kerning tables,
  missing fonts). Left/right/center lines tolerate this through the slack; a justified line that
  Word measures wider than `width` wraps.
- Only text runs, filled/stroked rectangles, straight lines and raster images are supported
  (no paths, gradients, rounded corners, clipping, rotation, dash patterns, opacity).
- SVG images are skipped (no raster fallback in the layout).
- Positions are exact in the file (EMU); Word's own PDF export quantizes to 0.24 pt.
- The text is editable, but it is laid out as independent text boxes: editing does not re-flow
  content across boxes or pages.
