using System.Text.Json.Serialization;

namespace OutlookComMcp.Models;

public sealed record MailSummary(
    string EmailId,
    string StoreId,
    string Subject,
    string SenderName,
    string SenderAddress,
    DateTimeOffset ReceivedAt,
    bool IsUnread,
    bool HasAttachments,
    [property: JsonIgnore(Condition = JsonIgnoreCondition.Never)]
    string? BodyPreview);

public sealed record MailDetail(
    string EmailId,
    string StoreId,
    string Subject,
    string SenderName,
    string SenderAddress,
    string To,
    string Cc,
    DateTimeOffset ReceivedAt,
    bool IsUnread,
    bool HasAttachments,
    string Body,
    bool BodyTruncated);

public sealed record CalendarEvent(
    string EventId,
    string StoreId,
    string Subject,
    DateTimeOffset StartsAt,
    DateTimeOffset EndsAt,
    string Location,
    string Organizer,
    bool IsAllDay,
    bool IsRecurring,
    string BusyStatus);

public sealed record ReplyDraft(
    string DraftId,
    string StoreId,
    string Subject,
    string To,
    string Cc,
    bool IsReplyAll,
    string Status);

public sealed record MailFolderInfo(
    string FolderId,
    string StoreId,
    string Name,
    string FolderPath,
    string? ParentFolderId,
    int UnreadItemCount,
    int TotalItemCount,
    bool HasChildren);

public sealed record EmailReadState(
    string EmailId,
    string StoreId,
    string Subject,
    DateTimeOffset ReceivedAt,
    bool IsRead,
    string Status);
