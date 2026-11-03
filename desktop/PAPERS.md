# ExamOffice Paper Authoring Guide

ExamOffice papers are defined as simple, standalone JSON manifests (`paper.json`). No code execution (`eval`, Python, shell) is allowed or required.

---

## Directory Structure

Built-in papers live inside `desktop/papers/`:
```
desktop/papers/
  w-1/
    paper.json
  w-2/
    paper.json
```

User-added papers are placed in `%LOCALAPPDATA%\ExamOffice\papers\`:
```
%LOCALAPPDATA%\ExamOffice\papers\
  my-paper-1/
    paper.json
```
ExamOffice automatically discovers all subdirectories in `%LOCALAPPDATA%\ExamOffice\papers\` that contain a `paper.json` file. You can also click the **"Open Papers Folder"** and **"Reload Papers"** buttons on the start screen.

---

## Schema Reference (`paper.json`)

```json
{
  "schema_version": 1,
  "id": "my-paper-1",
  "title": "Word Efficiency Test - Set 1",
  "duration_seconds": 600,
  "max_marks": 25,
  "pass_marks": 9.0,
  "pass_marks_scst": 7.75,
  "document": {
    "heading": "शीर्षक यहाँ लिखें",
    "paragraphs": [
      "पहला पैराग्राफ...",
      "दूसरा पैराग्राफ...",
      "तीसरा पैराग्राफ...",
      "चौथा पैराग्राफ..."
    ],
    "table": {
      "head": ["कॉलम 1", "कॉलम 2", "कॉलम 3"],
      "rows": [
        ["डेटा 1", "डेटा 2", "डेटा 3"],
        ["डेटा 4", "डेटा 5", ""]
      ]
    },
    "find_replace": {
      "find": "खोजने वाला शब्द",
      "replace": "बदलने वाला शब्द",
      "min_count": 2
    }
  },
  "tasks": [
    {
      "id": "t-01",
      "text": "पहले पैराग्राफ को Justify करें।",
      "marks": 1,
      "check": {
        "kind": "align",
        "para": 0,
        "val": "justify"
      },
      "hint_on_fail": "पहले पैराग्राफ का Alignment Justified नहीं मिला।"
    }
  ]
}
```

### Manifest Rules
- `schema_version`: Must be `1`.
- `id`: Unique identifier containing only letters, numbers, hyphens, and underscores.
- `duration_seconds`: Test countdown in seconds (must be ≥ 60; default 600 for 10 minutes).
- `max_marks`: Total possible marks (must be a positive integer).
- `tasks`: List of task objects. The sum of `marks` across all tasks **must exactly equal** `max_marks`.

---

## Supported Check Kinds

| `kind` | Parameters | Description |
|---|---|---|
| `align` | `para` (0-indexed), `val` ("justify" \| "center") | Paragraph text alignment |
| `bold` | `para` (0-indexed) | All runs in paragraph must be bold |
| `italic` | `para` (0-indexed) | All runs in paragraph must be italic |
| `underline` | `para` (0-indexed) | All runs in paragraph must be underlined |
| `linespacing` | `para` (0-indexed), `val` (float, e.g. `1.5` or `2.0`) | Paragraph line spacing multiplier |
| `indent` | `para` (0-indexed), `val` (float inches, e.g. `0.5`) | Left indent in inches |
| `first_line_indent` | `para` (0-indexed), `val` (float inches, e.g. `0.5`) | First-line indent in inches |
| `spaceafter` | `para` (0-indexed), `val` (float points, e.g. `12`) | Space after paragraph in points |
| `size` | `para` (0-indexed), `val` (float points, e.g. `14`) | Font size of all runs in paragraph |
| `color` | `para` (0-indexed), `val` (list of hex strings, e.g. `["FF0000"]`) | Font color of paragraph |
| `highlight` | `para` (0-indexed), `val` ("yellow") | Yellow highlight on paragraph |
| `heading_bold` | none | Document heading is bold |
| `heading_size` | `val` (float points, e.g. `16`) | Heading font size |
| `heading_color` | `val` (list of hex strings) | Heading font color |
| `heading_center_size` | `val` (float points, e.g. `16`) | Heading centered and font size set |
| `heading_bold_color` | `val` (list of hex strings) | Heading bold and font color set |
| `phrase_italic` | `para` (0-indexed), `val` (substring) | Specific phrase inside paragraph has italic formatting |
| `word_double_underline` | `para` (0-indexed), `val` (substring) | Specific word inside paragraph has double underline |
| `header_text` | `val` (string) | Header contains specified text |
| `footer_text` | `val` (string) | Footer contains specified text |
| `footer_pagenum` | none | Footer contains a PAGE number field or numeric text |
| `footer_text_bold` | `val` (string), `word` (bold substring) | Footer text contains `val` and `word` is bold |
| `find_replace` | none | Document no longer contains `document.find_replace.find` and contains `replace` at least `min_count` times |
| `orientation` | none | Page orientation is Landscape |
| `margins` | `val` (float inches, e.g. `0.5`) | Margins on all 4 sides match `val` |
| `orient_margins_narrow` | none | Orientation is Landscape and all margins are 0.5" |
| `page_color` | none | Page background color set (non-white, non-transparent) |
| `page_border` | none | Page borders applied to at least 3 edges |
| `page_break_before` | `para` (0-indexed) | Paragraph begins on a new page |
| `para_shading` | `para` (0-indexed) | Background shading applied to paragraph |
| `table_borders` | none | All table cells or outer borders have borders applied |
| `table_head_bold_shade` | none | Table header row is bold and shaded |
| `table_head_bold_center` | none | Table header row is bold and centered |
| `borders_and_cell` | `row` (0-indexed), `col` (0-indexed), `val` (string) | Table has borders and cell[row][col] matches text |
| `cell_text` | `row` (0-indexed), `col` (0-indexed), `val` (string) | Cell text matches `val` |
