using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Server;
using DocxMcp;
using DocxMcp.Persistence;
using DocxMcp.Tools;

var builder = Host.CreateApplicationBuilder(args);

// MCP requirement: all logging goes to stderr
builder.Logging.AddConsole(options =>
{
    options.LogToStandardErrorThreshold = LogLevel.Trace;
});

// Register persistence and session management
builder.Services.AddSingleton<SessionStore>();
builder.Services.AddSingleton<SessionManager>();
builder.Services.AddHostedService<SessionRestoreService>();

// Register MCP server with stdio transport and explicit tool types (AOT-safe)
builder.Services
    .AddMcpServer(options =>
    {
        options.ServerInfo = new()
        {
            Name = "docx-mcp",
            Version = "2.2.0"
        };
        options.Instructions = """
            DOCX-MCP: Word Document Manipulation Server

            ## Path Syntax

            Paths navigate document elements: /segment[selector]/segment[selector]/...

            ### Segments
            - /body — Document body (root)
            - /body/paragraph[N] — Paragraphs (alias: p)
            - /body/heading[level=N] — Headings by level
            - /body/table[N] — Tables
            - /body/table[N]/row[N] — Table rows
            - /body/table[N]/row[N]/cell[N] — Table cells
            - /body/paragraph[N]/run[N] — Text runs
            - /body/children/N — Positional insert point

            ### Selectors
            - [0], [-1] — By index (0-based, negative from end)
            - [id='1A2B3C4D'] — By stable ID (PREFERRED for modifications)
            - [text~='hello'] — Contains text (case-insensitive)
            - [text='Hello'] — Exact text match
            - [style='Heading1'] — By style name
            - [*] — All elements (wildcard)

            ## Best Practices

            1. **Query first**: Get element IDs before modifying
            2. **Use IDs for modifications**: /body/paragraph[id='ABC'] is stable
            3. **Use indexes for new documents**: Fine when building from scratch
            4. **Use dry_run**: Test patches with dry_run=true before applying
            5. **One replacement at a time**: replace_text defaults to max_count=1

            ## Response Format

            All patch operations return structured JSON:
            {
              "success": true,
              "applied": 1,
              "total": 1,
              "operations": [
                {"op": "add", "path": "...", "status": "success", "created_id": "1A2B3C4D"}
              ]
            }

            Operation-specific fields:
            - add: created_id
            - replace: replaced_id
            - remove: removed_id
            - move: moved_id, from
            - copy: source_id, copy_id
            - replace_text: matches_found, replacements_made
            - remove_column: column_index, rows_affected
            """;
    })
    .WithStdioServerTransport()
    // Document management
    .WithTools<DocumentTools>()
    // Query tools
    .WithTools<QueryTool>()
    .WithTools<CountTool>()
    .WithTools<ReadSectionTool>()
    .WithTools<ReadHeadingContentTool>()
    // Element operations (individual tools with focused documentation)
    .WithTools<ElementTools>()
    .WithTools<TextTools>()
    .WithTools<TableTools>()
    // Generic patch (multi-operation)
    .WithTools<PatchTool>()
    // Help and documentation
    .WithTools<PathHelpTool>()
    // Export, history, comments, styles
    .WithTools<ExportTools>()
    .WithTools<HistoryTools>()
    .WithTools<CommentTools>()
    .WithTools<StyleTools>();

await builder.Build().RunAsync();
