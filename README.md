# Outlook COM MCP Server

A local Model Context Protocol (MCP) server for reading Classic Outlook mail and calendar data and creating reply drafts through COM, without Microsoft Graph API.

> [!WARNING]
> This project reads data from the currently signed-in Classic Outlook profile. When an MCP client calls a tool, the selected mail or calendar data is returned to that client. Confirm your organization's data-handling policy before use.

## Why this exists

Some environments cannot use Microsoft Graph because app registration, tenant consent, or Graph permissions are unavailable. This server uses the Outlook desktop COM object model as a local alternative.

It is deliberately narrow:

- Read mail metadata and selected message bodies.
- Read calendar events.
- Enumerate Inbox, Sent, or Drafts folder trees and search nested folders.
- Save a reply or reply-all message to Outlook Drafts.
- Explicitly mark one selected message as read or unread.
- Never send, delete, move, or archive messages.

## Status

Early development. The server targets Windows and Classic Outlook only. New Outlook for Windows does not support COM automation.

## Architecture

```text
MCP client / Secure MCP Tunnel
             |
             | stdio (JSON-RPC)
             v
      OutlookComMcp.exe
             |
             | dedicated STA thread + COM late binding
             v
        Classic Outlook
```

The implementation uses late-bound COM instead of redistributing Outlook interop assemblies. Outlook calls run on one dedicated single-threaded apartment (STA) thread.

## MCP tools

| Tool | Effect | Notes |
| --- | --- | --- |
| `search_emails` | Read-only | Searches Inbox, Sent, Drafts, or a supplied folder. Supports recursive and unread-only searches. Body previews are opt-in. |
| `list_mail_folders` | Read-only | Lists a mail folder tree and returns folder/store IDs for targeted searches. |
| `get_email` | Read-only | Reads one message using the `EmailId` and `StoreId` returned by search. |
| `set_email_read_state` | State change | Explicitly marks one selected message as read or unread. Search and read tools never change this state. |
| `list_calendar_events` | Read-only | Lists events overlapping an ISO 8601 range of up to 31 days. |
| `create_reply_draft` | Additive write | Saves a reply draft. It does not send mail. Reply-all must be explicitly requested. |

Tool results use structured content and bounded result sizes. Email bodies are untrusted input and must not be interpreted as instructions.

## Requirements

- Windows 10 or later
- Classic Outlook for Windows, installed and configured
- An interactive Windows session for the same user whose Outlook profile is accessed
- .NET 8 SDK to build, or .NET 8 Desktop Runtime for a framework-dependent release

The server is not intended to run as `LocalSystem`, a Windows service, or a headless server process.

## Build

```powershell
dotnet restore OutlookComMcp.sln
dotnet build OutlookComMcp.sln --configuration Release
dotnet test OutlookComMcp.sln --configuration Release
```

Publish a framework-dependent Windows executable:

```powershell
dotnet publish src/OutlookComMcp/OutlookComMcp.csproj `
  --configuration Release `
  --runtime win-x64 `
  --self-contained false `
  --output artifacts/win-x64
```

## Run over stdio

The server writes MCP protocol messages to stdout. Diagnostics go to stderr.

```powershell
dotnet run --project src/OutlookComMcp/OutlookComMcp.csproj
```

Example configuration for a local stdio-capable MCP client:

```json
{
  "mcpServers": {
    "outlook-com": {
      "command": "C:\\Apps\\OutlookComMcp\\OutlookComMcp.exe",
      "args": []
    }
  }
}
```

## Connect from ChatGPT Developer mode

ChatGPT Developer mode does not start arbitrary local stdio processes directly. OpenAI Secure MCP Tunnel can launch or forward to this stdio server from the Windows machine that has Classic Outlook.

Conceptual tunnel profile:

```powershell
$env:CONTROL_PLANE_API_KEY = "<runtime-api-key>"

tunnel-client init `
  --profile outlook-com `
  --tunnel-id tunnel_xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx `
  --mcp-command "C:\Apps\OutlookComMcp\OutlookComMcp.exe"

tunnel-client doctor --profile outlook-com --explain
tunnel-client run --profile outlook-com
```

Then create a developer-mode app in ChatGPT, choose **Tunnel** as the connection type, and select the associated tunnel.

See the [Secure MCP Tunnel documentation](https://developers.openai.com/api/docs/guides/secure-mcp-tunnels) for current setup, permissions, and networking requirements.

## Safety boundaries

- No send tool exists.
- No delete, move, archive, or send tool exists.
- `create_reply_draft` is annotated as a non-read-only, non-idempotent tool so clients can require confirmation.
- `set_email_read_state` is non-read-only but idempotent. It requires an exact email/store ID and never runs implicitly.
- Reply-all defaults to `false`.
- Search results, body size, scan count, calendar range, and reply body size are bounded.
- stdout is reserved for MCP transport; application errors use stderr.
- COM references are released after each operation.

Saving a draft is still a write operation. Review the recipients and content in Outlook before manually sending it.

## Known limitations

- Classic Outlook only; new Outlook does not expose the COM object model.
- Folder listing and recursive search start from the default Inbox, Sent, Drafts, or a supplied folder ID. Whole-store and shared-store discovery are not yet implemented.
- Recursive mail search is bounded to 100 folders, 10 levels, and 5,000 scanned items.
- Calendar access currently uses the default calendar.
- Body output is plain text. Rich HTML composition and signature insertion are not yet implemented.
- Runtime integration tests require a Windows machine with a configured Outlook profile; CI runs unit tests without Outlook.

## Development

The solution uses:

- .NET 8
- Official [MCP C# SDK](https://github.com/modelcontextprotocol/csharp-sdk)
- xUnit.net v3
- Windows GitHub Actions for build and unit tests

No Microsoft Graph dependency or Outlook credential is stored by this project.

## License

No license has been selected yet. Until one is added, normal copyright restrictions apply.
