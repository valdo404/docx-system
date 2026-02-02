using DocxMcp.Persistence;
using Microsoft.Extensions.Logging.Abstractions;

namespace DocxMcp.Tests;

internal static class TestHelpers
{
    /// <summary>
    /// Create a SessionManager backed by a temporary directory for testing.
    /// Each call creates a unique temp directory so tests don't interfere.
    /// </summary>
    public static SessionManager CreateSessionManager()
    {
        var tempDir = Path.Combine(Path.GetTempPath(), "docx-mcp-tests", Guid.NewGuid().ToString("N"));
        var store = new SessionStore(NullLogger<SessionStore>.Instance, tempDir);
        return new SessionManager(store, NullLogger<SessionManager>.Instance);
    }
}
