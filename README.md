# docx-mcp

A Model Context Protocol (MCP) server for Microsoft Word DOCX manipulation, built with .NET 10 and NativeAOT. Single binary, no runtime required.

## Why docx-mcp?

Most document tools treat Word files as opaque blobs — open, edit, save, hope for the best. docx-mcp treats them as structured data with full version control.

**Structured Query Language for Word** — Navigate documents with typed paths like `/body/heading[level=1]` or `/body/paragraph[text~='budget']`. No more guessing at element indices.

**Time Travel (Undo/Redo)** — Every edit is recorded in a write-ahead log. Undo, redo, or jump to any point in the editing history. Checkpoints every N operations keep rebuilds fast. Branch off from any past state — the future timeline is automatically discarded.

**Session Persistence** — Sessions survive server restarts. The WAL, checkpoints, and cursor position are durably stored. Pick up exactly where you left off, even across Docker container restarts.

**Patch-Based Editing** — RFC 6902 JSON Patch adapted for OOXML. Add headings, tables, images, hyperlinks, lists. Replace text while preserving run-level formatting. Move, copy, remove elements. Up to 10 operations per call, batched atomically.

**NativeAOT** — Compiles to a single ~28MB binary with no .NET runtime dependency. Sub-millisecond startup. Ships as a minimal Docker image.

## Quick Start

### Docker

```bash
docker pull valdo404/docx-mcp
```

```json
{
  "mcpServers": {
    "docx": {
      "command": "docker",
      "args": ["run", "-i", "--rm",
        "-v", "/path/to/documents:/data",
        "-v", "docx-sessions:/home/app/.docx-mcp/sessions",
        "valdo404/docx-mcp"]
    }
  }
}
```

### Native Binary

```bash
# Build the NativeAOT binary (requires .NET 10 SDK)
./publish.sh

# Binary output: dist/macos-arm64/docx-mcp (~28MB)
```

## AI Tool Integration

### Claude Desktop

**macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`

Using Docker (recommended):
```json
{
  "mcpServers": {
    "docx": {
      "command": "docker",
      "args": ["run", "-i", "--rm",
        "-v", "/Users/you/Documents:/data",
        "-v", "docx-sessions:/home/app/.docx-mcp/sessions",
        "valdo404/docx-mcp"]
    }
  }
}
```

Using native binary:
```json
{
  "mcpServers": {
    "docx": {
      "command": "/absolute/path/to/docx-mcp/dist/macos-arm64/docx-mcp",
      "args": []
    }
  }
}
```

### Cursor

Using Docker:
```json
{
  "mcp": {
    "servers": {
      "docx": {
        "command": "docker",
        "args": ["run", "-i", "--rm",
          "-v", "/Users/you/Documents:/data",
          "-v", "docx-sessions:/home/app/.docx-mcp/sessions",
          "valdo404/docx-mcp"]
      }
    }
  }
}
```

Using native binary:
```json
{
  "mcp": {
    "servers": {
      "docx": {
        "command": "/absolute/path/to/docx-mcp/dist/macos-arm64/docx-mcp",
        "args": []
      }
    }
  }
}
```

## Tools

The server exposes 18 tools over MCP stdio transport:

### Document Management

| Tool | Description |
|------|-------------|
| `document_open` | Open a .docx file or create a new empty document. Returns a session ID. |
| `document_save` | Save document to disk (original path or new path). |
| `document_close` | Close session and release resources. |
| `document_list` | List all open document sessions. |

### Query

| Tool | Description |
|------|-------------|
| `query` | Read any part of a document using typed paths. Returns JSON, text, or summary. |

**Path examples:**

```
/body                           — document structure summary
/body/paragraph[0]              — first paragraph
/body/paragraph[*]              — all paragraphs
/body/table[0]                  — first table
/body/heading[*]                — all headings
/body/heading[level=1]          — level-1 headings only
/body/paragraph[text~='hello']  — paragraphs containing 'hello'
/metadata                       — document properties
/styles                         — style definitions
```

### Editing

| Tool | Description |
|------|-------------|
| `apply_patch` | Modify documents using JSON patches (RFC 6902 adapted for OOXML). |

**Operations:** `add`, `replace`, `remove`, `move`, `copy`

**Element types:** `paragraph`, `heading`, `table`, `image`, `hyperlink`, `page_break`, `list`

**Example patches:**

```json
[
  {"op": "add", "path": "/body/children/0", "value": {"type": "heading", "level": 1, "text": "Title"}},
  {"op": "add", "path": "/body/children/1", "value": {"type": "paragraph", "text": "Hello world.", "style": {"bold": true}}},
  {"op": "add", "path": "/body/children/2", "value": {"type": "table", "headers": ["Name", "Value"], "rows": [["foo", "bar"]]}},
  {"op": "remove", "path": "/body/paragraph[0]"}
]
```

### History & Time Travel

| Tool | Description |
|------|-------------|
| `document_undo` | Undo N steps. Rebuilds from the nearest checkpoint. |
| `document_redo` | Redo N steps. Replays patches forward (no rebuild needed). |
| `document_history` | List all WAL entries with timestamps, descriptions, and current position. |
| `document_jump_to` | Jump to any position in the editing timeline. |

Every `apply_patch` call is recorded with a timestamp and auto-generated description. Undo rebuilds the document from the nearest checkpoint (snapshots taken every 10 edits by default, configurable via `DOCX_MCP_CHECKPOINT_INTERVAL`). Redo replays patches forward on the current DOM — no rebuild overhead.

Applying a new patch after an undo discards the future timeline and starts a new branch, just like typing after undo in a text editor.

### Export

| Tool | Description |
|------|-------------|
| `export_pdf` | Export to PDF via LibreOffice CLI (requires LibreOffice installed). |
| `export_html` | Export to HTML. |
| `export_markdown` | Export to Markdown. |

### Additional Tools

| Tool | Description |
|------|-------------|
| `read_section` | Read content under a specific heading (section-based navigation). |
| `read_heading_content` | Read content between two headings. |
| `document_count` | Count elements by type (paragraphs, tables, headings, etc.). |
| `document_snapshot` | Compact the WAL into a single baseline. Optionally discard redo history. |

## Building

### Prerequisites

- [.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0)
- macOS: Homebrew (`openssl`, `brotli` — needed for NativeAOT linking)

### Build

```bash
# Build for current platform
./publish.sh

# Build for a specific target
./publish.sh macos-arm64
./publish.sh linux-x64

# Build all targets
./publish.sh all
```

**Supported targets:** `macos-arm64`, `macos-x64`, `linux-x64`, `linux-arm64`, `windows-x64`, `windows-arm64`

Output goes to `dist/<target>/docx-mcp`.

### Tests

```bash
# Unit tests (257 tests)
dotnet test tests/DocxMcp.Tests/

# Integration tests (requires mcptools: brew install mcptools)
./test-mcp.sh

# Integration test with a real document
./test-mcp.sh ~/Documents/somefile.docx
```

## Architecture

```
src/DocxMcp/
  Program.cs              — MCP server setup (stdio transport)
  SessionManager.cs       — Document session lifecycle + undo/redo
  DocxSession.cs          — Single document wrapper
  Tools/
    DocumentTools.cs      — open / save / close / list / snapshot
    QueryTool.cs          — typed path queries
    PatchTool.cs          — JSON patch operations
    HistoryTools.cs       — undo / redo / history / jump_to
    ExportTools.cs        — PDF / HTML / Markdown export
  Persistence/
    SessionStore.cs       — WAL, checkpoint, and index I/O
    MappedWal.cs          — Memory-mapped write-ahead log with random access
    SessionIndex.cs       — Session metadata (cursor, checkpoints)
    WalEntry.cs           — WAL entry model (patches, timestamp, description)
  Paths/
    DocxPath.cs           — Path model (/body/paragraph[0])
    PathParser.cs         — Path string parser
    PathResolver.cs       — Path-to-OOXML element resolution
    Selector.cs           — Filter predicates (index, wildcard, attribute, text search)
  Helpers/
    ElementFactory.cs     — JSON → OOXML element creation
    OpenXmlExtensions.cs  — Extension methods for OpenXml types
```

**Stack:** .NET 10, Open XML SDK 3.2, MCP C# SDK 0.7.0, NativeAOT

## License

MIT
