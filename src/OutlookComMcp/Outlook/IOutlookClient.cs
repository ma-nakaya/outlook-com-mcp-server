using OutlookComMcp.Models;

namespace OutlookComMcp.Outlook;

public interface IOutlookClient
{
    Task<IReadOnlyList<MailSummary>> SearchEmailsAsync(
        string folder,
        string? query,
        int daysBack,
        int maxResults,
        bool includeBodyPreview,
        string? folderId,
        string? storeId,
        bool includeSubfolders,
        bool unreadOnly,
        CancellationToken cancellationToken);

    Task<IReadOnlyList<MailFolderInfo>> ListMailFoldersAsync(
        string folder,
        string? parentFolderId,
        string? storeId,
        bool recursive,
        int maxResults,
        int maxDepth,
        CancellationToken cancellationToken);

    Task<MailDetail> GetEmailAsync(
        string emailId,
        string storeId,
        int maxBodyCharacters,
        CancellationToken cancellationToken);

    Task<IReadOnlyList<CalendarEvent>> ListCalendarEventsAsync(
        DateTimeOffset startsAfter,
        DateTimeOffset endsBefore,
        int maxResults,
        CancellationToken cancellationToken);

    Task<EmailReadState> SetEmailReadStateAsync(
        string emailId,
        string storeId,
        bool isRead,
        CancellationToken cancellationToken);

    Task<ReplyDraft> CreateReplyDraftAsync(
        string emailId,
        string storeId,
        string body,
        bool replyAll,
        CancellationToken cancellationToken);
}

