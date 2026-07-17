using System.Globalization;
using OutlookComMcp.Models;

namespace OutlookComMcp.Outlook;

public sealed class OutlookComClient(StaDispatcher dispatcher) : IOutlookClient
{
    private const int MailItemClass = 43;
    private const int AppointmentItemClass = 26;
    private const int InboxFolder = 6;
    private const int SentMailFolder = 5;
    private const int DraftsFolder = 16;
    private const int MaximumScannedItems = 5_000;
    private const int BodyPreviewCharacters = 500;

    public Task<IReadOnlyList<MailSummary>> SearchEmailsAsync(
        string folder,
        string? query,
        int daysBack,
        int maxResults,
        bool includeBodyPreview,
        CancellationToken cancellationToken) =>
        dispatcher.InvokeAsync<IReadOnlyList<MailSummary>>(
            () => SearchEmails(folder, query, daysBack, maxResults, includeBodyPreview),
            cancellationToken);

    public Task<MailDetail> GetEmailAsync(
        string emailId,
        string storeId,
        int maxBodyCharacters,
        CancellationToken cancellationToken) =>
        dispatcher.InvokeAsync(
            () => GetEmail(emailId, storeId, maxBodyCharacters),
            cancellationToken);

    public Task<IReadOnlyList<CalendarEvent>> ListCalendarEventsAsync(
        DateTimeOffset startsAfter,
        DateTimeOffset endsBefore,
        int maxResults,
        CancellationToken cancellationToken) =>
        dispatcher.InvokeAsync<IReadOnlyList<CalendarEvent>>(
            () => ListCalendarEvents(startsAfter, endsBefore, maxResults),
            cancellationToken);

    public Task<ReplyDraft> CreateReplyDraftAsync(
        string emailId,
        string storeId,
        string body,
        bool replyAll,
        CancellationToken cancellationToken) =>
        dispatcher.InvokeAsync(
            () => CreateReplyDraft(emailId, storeId, body, replyAll),
            cancellationToken);

    private static IReadOnlyList<MailSummary> SearchEmails(
        string folderName,
        string? query,
        int daysBack,
        int maxResults,
        bool includeBodyPreview)
    {
        object? application = null;
        object? session = null;
        object? folder = null;
        object? items = null;

        try
        {
            (application, session) = OpenSession();
            dynamic outlookSession = session;
            folder = outlookSession.GetDefaultFolder(GetFolderId(folderName));
            dynamic outlookFolder = folder;
            string storeId = Convert.ToString(outlookFolder.StoreID, CultureInfo.InvariantCulture) ?? string.Empty;

            items = outlookFolder.Items;
            dynamic outlookItems = items;
            outlookItems.Sort("[ReceivedTime]", true);

            DateTime cutoff = DateTime.Now.AddDays(-daysBack);
            string normalizedQuery = query?.Trim() ?? string.Empty;
            List<MailSummary> results = [];
            int itemCount = Math.Min(Convert.ToInt32(outlookItems.Count, CultureInfo.InvariantCulture), MaximumScannedItems);

            for (int index = 1; index <= itemCount && results.Count < maxResults; index++)
            {
                object? item = null;
                try
                {
                    item = outlookItems[index];
                    dynamic mail = item;
                    if (Convert.ToInt32(mail.Class, CultureInfo.InvariantCulture) != MailItemClass)
                    {
                        continue;
                    }

                    DateTime receivedAt = Convert.ToDateTime(mail.ReceivedTime, CultureInfo.CurrentCulture);
                    if (receivedAt < cutoff)
                    {
                        break;
                    }

                    string subject = Convert.ToString(mail.Subject, CultureInfo.CurrentCulture) ?? string.Empty;
                    string senderName = Convert.ToString(mail.SenderName, CultureInfo.CurrentCulture) ?? string.Empty;
                    string senderAddress = ResolveSenderAddress(mail);
                    if (!Matches(normalizedQuery, subject, senderName, senderAddress))
                    {
                        continue;
                    }

                    string? preview = null;
                    if (includeBodyPreview)
                    {
                        preview = Truncate(Convert.ToString(mail.Body, CultureInfo.CurrentCulture) ?? string.Empty, BodyPreviewCharacters);
                    }

                    results.Add(new(
                        Convert.ToString(mail.EntryID, CultureInfo.InvariantCulture) ?? string.Empty,
                        storeId,
                        subject,
                        senderName,
                        senderAddress,
                        ToDateTimeOffset(receivedAt),
                        Convert.ToBoolean(mail.UnRead, CultureInfo.InvariantCulture),
                        HasAttachments(mail),
                        preview));
                }
                finally
                {
                    ComObject.Release(item);
                }
            }

            return results;
        }
        finally
        {
            ComObject.Release(items);
            ComObject.Release(folder);
            ComObject.Release(session);
            ComObject.Release(application);
        }
    }

    private static MailDetail GetEmail(string emailId, string storeId, int maxBodyCharacters)
    {
        object? application = null;
        object? session = null;
        object? item = null;

        try
        {
            (application, session) = OpenSession();
            dynamic outlookSession = session;
            item = outlookSession.GetItemFromID(emailId, storeId);
            dynamic mail = item;
            EnsureMailItem(mail);

            string body = Convert.ToString(mail.Body, CultureInfo.CurrentCulture) ?? string.Empty;
            bool truncated = body.Length > maxBodyCharacters;
            if (truncated)
            {
                body = body[..maxBodyCharacters];
            }

            DateTime receivedAt = Convert.ToDateTime(mail.ReceivedTime, CultureInfo.CurrentCulture);
            return new(
                Convert.ToString(mail.EntryID, CultureInfo.InvariantCulture) ?? emailId,
                storeId,
                Convert.ToString(mail.Subject, CultureInfo.CurrentCulture) ?? string.Empty,
                Convert.ToString(mail.SenderName, CultureInfo.CurrentCulture) ?? string.Empty,
                ResolveSenderAddress(mail),
                Convert.ToString(mail.To, CultureInfo.CurrentCulture) ?? string.Empty,
                Convert.ToString(mail.CC, CultureInfo.CurrentCulture) ?? string.Empty,
                ToDateTimeOffset(receivedAt),
                Convert.ToBoolean(mail.UnRead, CultureInfo.InvariantCulture),
                HasAttachments(mail),
                body,
                truncated);
        }
        finally
        {
            ComObject.Release(item);
            ComObject.Release(session);
            ComObject.Release(application);
        }
    }

    private static IReadOnlyList<CalendarEvent> ListCalendarEvents(
        DateTimeOffset startsAfter,
        DateTimeOffset endsBefore,
        int maxResults)
    {
        object? application = null;
        object? session = null;
        object? folder = null;
        object? items = null;
        object? current = null;

        try
        {
            (application, session) = OpenSession();
            dynamic outlookSession = session;
            folder = outlookSession.GetDefaultFolder(9);
            dynamic outlookFolder = folder;
            string storeId = Convert.ToString(outlookFolder.StoreID, CultureInfo.InvariantCulture) ?? string.Empty;

            items = outlookFolder.Items;
            dynamic outlookItems = items;
            outlookItems.Sort("[Start]", false);
            outlookItems.IncludeRecurrences = true;

            DateTime localStart = startsAfter.LocalDateTime;
            DateTime localEnd = endsBefore.LocalDateTime;
            List<CalendarEvent> results = [];
            current = outlookItems.GetFirst();

            for (int scanned = 0; current is not null && scanned < MaximumScannedItems && results.Count < maxResults; scanned++)
            {
                object? next = null;
                try
                {
                    dynamic appointment = current;
                    next = outlookItems.GetNext();
                    if (Convert.ToInt32(appointment.Class, CultureInfo.InvariantCulture) != AppointmentItemClass)
                    {
                        continue;
                    }

                    DateTime start = Convert.ToDateTime(appointment.Start, CultureInfo.CurrentCulture);
                    DateTime end = Convert.ToDateTime(appointment.End, CultureInfo.CurrentCulture);
                    if (start > localEnd)
                    {
                        break;
                    }

                    if (end < localStart)
                    {
                        continue;
                    }

                    results.Add(new(
                        Convert.ToString(appointment.EntryID, CultureInfo.InvariantCulture) ?? string.Empty,
                        storeId,
                        Convert.ToString(appointment.Subject, CultureInfo.CurrentCulture) ?? string.Empty,
                        ToDateTimeOffset(start),
                        ToDateTimeOffset(end),
                        Convert.ToString(appointment.Location, CultureInfo.CurrentCulture) ?? string.Empty,
                        Convert.ToString(appointment.Organizer, CultureInfo.CurrentCulture) ?? string.Empty,
                        Convert.ToBoolean(appointment.AllDayEvent, CultureInfo.InvariantCulture),
                        Convert.ToBoolean(appointment.IsRecurring, CultureInfo.InvariantCulture),
                        BusyStatusName(Convert.ToInt32(appointment.BusyStatus, CultureInfo.InvariantCulture))));
                }
                finally
                {
                    ComObject.Release(current);
                    current = next;
                }
            }

            return results;
        }
        finally
        {
            ComObject.Release(current);
            ComObject.Release(items);
            ComObject.Release(folder);
            ComObject.Release(session);
            ComObject.Release(application);
        }
    }

    private static ReplyDraft CreateReplyDraft(string emailId, string storeId, string body, bool replyAll)
    {
        object? application = null;
        object? session = null;
        object? original = null;
        object? reply = null;
        object? parent = null;

        try
        {
            (application, session) = OpenSession();
            dynamic outlookSession = session;
            original = outlookSession.GetItemFromID(emailId, storeId);
            dynamic originalMail = original;
            EnsureMailItem(originalMail);

            reply = replyAll ? originalMail.ReplyAll() : originalMail.Reply();
            dynamic replyMail = reply;
            string existingBody = Convert.ToString(replyMail.Body, CultureInfo.CurrentCulture) ?? string.Empty;
            replyMail.Body = body.Trim() + Environment.NewLine + Environment.NewLine + existingBody;
            replyMail.Save();

            parent = replyMail.Parent;
            dynamic parentFolder = parent;
            return new(
                Convert.ToString(replyMail.EntryID, CultureInfo.InvariantCulture) ?? string.Empty,
                Convert.ToString(parentFolder.StoreID, CultureInfo.InvariantCulture) ?? string.Empty,
                Convert.ToString(replyMail.Subject, CultureInfo.CurrentCulture) ?? string.Empty,
                Convert.ToString(replyMail.To, CultureInfo.CurrentCulture) ?? string.Empty,
                Convert.ToString(replyMail.CC, CultureInfo.CurrentCulture) ?? string.Empty,
                replyAll,
                "saved_to_drafts");
        }
        finally
        {
            ComObject.Release(parent);
            ComObject.Release(reply);
            ComObject.Release(original);
            ComObject.Release(session);
            ComObject.Release(application);
        }
    }

    private static (object Application, object Session) OpenSession()
    {
        Type outlookType = Type.GetTypeFromProgID("Outlook.Application")
            ?? throw new InvalidOperationException(
                "Classic Outlook is not installed or Outlook COM automation is unavailable.");

        object application = Activator.CreateInstance(outlookType)
            ?? throw new InvalidOperationException("Could not start Classic Outlook.");

        try
        {
            dynamic outlookApplication = application;
            object session = outlookApplication.GetNamespace("MAPI");
            return (application, session);
        }
        catch
        {
            ComObject.Release(application);
            throw;
        }
    }

    private static void EnsureMailItem(dynamic item)
    {
        if (Convert.ToInt32(item.Class, CultureInfo.InvariantCulture) != MailItemClass)
        {
            throw new InvalidOperationException("The selected Outlook item is not an email message.");
        }
    }

    private static int GetFolderId(string folder) => folder.Trim().ToLowerInvariant() switch
    {
        "inbox" => InboxFolder,
        "sent" or "sent_mail" => SentMailFolder,
        "drafts" => DraftsFolder,
        _ => throw new ArgumentOutOfRangeException(nameof(folder), "Supported folders: inbox, sent, drafts."),
    };

    private static bool Matches(string query, params string[] values) =>
        string.IsNullOrWhiteSpace(query) ||
        values.Any(value => value.Contains(query, StringComparison.CurrentCultureIgnoreCase));

    private static string ResolveSenderAddress(dynamic mail)
    {
        string address = Convert.ToString(mail.SenderEmailAddress, CultureInfo.CurrentCulture) ?? string.Empty;
        string addressType = Convert.ToString(mail.SenderEmailType, CultureInfo.InvariantCulture) ?? string.Empty;
        if (!string.Equals(addressType, "EX", StringComparison.OrdinalIgnoreCase))
        {
            return address;
        }

        object? sender = null;
        object? exchangeUser = null;
        try
        {
            sender = mail.Sender;
            if (sender is null)
            {
                return address;
            }

            dynamic senderObject = sender;
            exchangeUser = senderObject.GetExchangeUser();
            if (exchangeUser is null)
            {
                return address;
            }

            dynamic user = exchangeUser;
            return Convert.ToString(user.PrimarySmtpAddress, CultureInfo.InvariantCulture) ?? address;
        }
        catch
        {
            return address;
        }
        finally
        {
            ComObject.Release(exchangeUser);
            ComObject.Release(sender);
        }
    }

    private static bool HasAttachments(dynamic mail)
    {
        object? attachments = null;
        try
        {
            attachments = mail.Attachments;
            dynamic attachmentCollection = attachments;
            return Convert.ToInt32(attachmentCollection.Count, CultureInfo.InvariantCulture) > 0;
        }
        finally
        {
            ComObject.Release(attachments);
        }
    }

    private static DateTimeOffset ToDateTimeOffset(DateTime value)
    {
        DateTime local = DateTime.SpecifyKind(value, DateTimeKind.Unspecified);
        return new(local, TimeZoneInfo.Local.GetUtcOffset(local));
    }

    private static string Truncate(string value, int maxCharacters) =>
        value.Length <= maxCharacters ? value : value[..maxCharacters];

    private static string BusyStatusName(int value) => value switch
    {
        0 => "free",
        1 => "tentative",
        2 => "busy",
        3 => "out_of_office",
        4 => "working_elsewhere",
        _ => "unknown",
    };
}
