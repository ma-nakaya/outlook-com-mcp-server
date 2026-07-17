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

    Task<ReplyDraft> CreateReplyDraftAsync(
        string emailId,
        string storeId,
        string body,
        bool replyAll,
        CancellationToken cancellationToken);
}

