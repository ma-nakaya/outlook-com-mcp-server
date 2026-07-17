using System.Text.Json;
using System.Text.Json.Serialization;
using OutlookComMcp.Models;
using OutlookComMcp.Outlook;
using OutlookComMcp.Tools;
using Xunit;

namespace OutlookComMcp.Tests;

public sealed class OutlookToolsTests
{
    private static readonly JsonSerializerOptions NullOmittingJsonOptions = new()
    {
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
    };

    [Fact]
    public void MailSummarySerializesNullBodyPreview()
    {
        MailSummary summary = new(
            "email-id",
            "store-id",
            "Subject",
            "Sender",
            "sender@example.com",
            new DateTimeOffset(2026, 7, 17, 9, 0, 0, TimeSpan.FromHours(9)),
            false,
            false,
            null);

        string json = JsonSerializer.Serialize(summary, NullOmittingJsonOptions);

        Assert.Contains("\"bodyPreview\":null", json, StringComparison.Ordinal);
    }

    [Fact]
    public async Task SearchEmailsPassesValidatedArgumentsToClient()
    {
        FakeOutlookClient client = new();
        OutlookTools tools = new(client);

        await tools.SearchEmailsAsync(
            "inbox",
            "status",
            14,
            10,
            true,
            includeSubfolders: true,
            unreadOnly: true,
            cancellationToken: TestContext.Current.CancellationToken);

        Assert.Equal(
            ("inbox", "status", 14, 10, true, null, null, true, true),
            client.SearchArguments);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(51)]
    public async Task SearchEmailsRejectsInvalidMaximum(int maxResults)
    {
        OutlookTools tools = new(new FakeOutlookClient());

        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(
            () => tools.SearchEmailsAsync(
                "inbox",
                null,
                30,
                maxResults,
                false,
                cancellationToken: TestContext.Current.CancellationToken));
    }

    [Fact]
    public async Task SearchEmailsRequiresStoreForCustomFolder()
    {
        OutlookTools tools = new(new FakeOutlookClient());

        await Assert.ThrowsAsync<ArgumentException>(
            () => tools.SearchEmailsAsync(
                folderId: "folder-id",
                cancellationToken: TestContext.Current.CancellationToken));
    }

    [Fact]
    public async Task ListMailFoldersPassesValidatedArgumentsToClient()
    {
        FakeOutlookClient client = new();
        OutlookTools tools = new(client);

        await tools.ListMailFoldersAsync(
            "inbox",
            recursive: true,
            maxResults: 75,
            maxDepth: 4,
            cancellationToken: TestContext.Current.CancellationToken);

        Assert.Equal(("inbox", null, null, true, 75, 4), client.FolderArguments);
    }

    [Fact]
    public async Task SetEmailReadStateIsExplicitAndIdempotent()
    {
        FakeOutlookClient client = new();
        OutlookTools tools = new(client);

        EmailReadState result = await tools.SetEmailReadStateAsync(
            "email-id",
            "store-id",
            true,
            TestContext.Current.CancellationToken);

        Assert.Equal(("email-id", "store-id", true), client.ReadStateArguments);
        Assert.True(result.IsRead);
    }

    [Fact]
    public async Task CalendarRejectsRangesLongerThanThirtyOneDays()
    {
        OutlookTools tools = new(new FakeOutlookClient());

        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(
            () => tools.ListCalendarEventsAsync(
                "2026-07-01T00:00:00+09:00",
                "2026-08-02T00:00:01+09:00",
                50,
                TestContext.Current.CancellationToken));
    }

    [Fact]
    public async Task CreateReplyDraftDoesNotSendAndDefaultsToReplyOnly()
    {
        FakeOutlookClient client = new();
        OutlookTools tools = new(client);

        ReplyDraft result = await tools.CreateReplyDraftAsync(
            "email-id",
            "store-id",
            "Thanks for the update.",
            cancellationToken: TestContext.Current.CancellationToken);

        Assert.Equal(("email-id", "store-id", "Thanks for the update.", false), client.DraftArguments);
        Assert.Equal("saved_to_drafts", result.Status);
    }

    [Fact]
    public async Task CreateReplyDraftRejectsEmptyBody()
    {
        OutlookTools tools = new(new FakeOutlookClient());

        await Assert.ThrowsAsync<ArgumentException>(
            () => tools.CreateReplyDraftAsync(
                "email-id",
                "store-id",
                "  ",
                cancellationToken: TestContext.Current.CancellationToken));
    }

    private sealed class FakeOutlookClient : IOutlookClient
    {
        public (
            string Folder,
            string? Query,
            int DaysBack,
            int MaxResults,
            bool IncludeBodyPreview,
            string? FolderId,
            string? StoreId,
            bool IncludeSubfolders,
            bool UnreadOnly)? SearchArguments { get; private set; }

        public (string Folder, string? ParentFolderId, string? StoreId, bool Recursive, int MaxResults, int MaxDepth)? FolderArguments { get; private set; }

        public (string EmailId, string StoreId, bool IsRead)? ReadStateArguments { get; private set; }

        public (string EmailId, string StoreId, string Body, bool ReplyAll)? DraftArguments { get; private set; }

        public Task<IReadOnlyList<MailSummary>> SearchEmailsAsync(
            string folder,
            string? query,
            int daysBack,
            int maxResults,
            bool includeBodyPreview,
            string? folderId,
            string? storeId,
            bool includeSubfolders,
            bool unreadOnly,
            CancellationToken cancellationToken)
        {
            SearchArguments = (
                folder,
                query,
                daysBack,
                maxResults,
                includeBodyPreview,
                folderId,
                storeId,
                includeSubfolders,
                unreadOnly);
            return Task.FromResult<IReadOnlyList<MailSummary>>([]);
        }

        public Task<IReadOnlyList<MailFolderInfo>> ListMailFoldersAsync(
            string folder,
            string? parentFolderId,
            string? storeId,
            bool recursive,
            int maxResults,
            int maxDepth,
            CancellationToken cancellationToken)
        {
            FolderArguments = (folder, parentFolderId, storeId, recursive, maxResults, maxDepth);
            return Task.FromResult<IReadOnlyList<MailFolderInfo>>([]);
        }

        public Task<MailDetail> GetEmailAsync(
            string emailId,
            string storeId,
            int maxBodyCharacters,
            CancellationToken cancellationToken) => throw new NotImplementedException();

        public Task<EmailReadState> SetEmailReadStateAsync(
            string emailId,
            string storeId,
            bool isRead,
            CancellationToken cancellationToken)
        {
            ReadStateArguments = (emailId, storeId, isRead);
            return Task.FromResult(new EmailReadState(
                emailId,
                storeId,
                "Subject",
                new DateTimeOffset(2026, 7, 17, 9, 0, 0, TimeSpan.FromHours(9)),
                isRead,
                "read_state_updated"));
        }

        public Task<IReadOnlyList<CalendarEvent>> ListCalendarEventsAsync(
            DateTimeOffset startsAfter,
            DateTimeOffset endsBefore,
            int maxResults,
            CancellationToken cancellationToken) =>
            Task.FromResult<IReadOnlyList<CalendarEvent>>([]);

        public Task<ReplyDraft> CreateReplyDraftAsync(
            string emailId,
            string storeId,
            string body,
            bool replyAll,
            CancellationToken cancellationToken)
        {
            DraftArguments = (emailId, storeId, body, replyAll);
            return Task.FromResult(new ReplyDraft(
                "draft-id",
                "store-id",
                "RE: Subject",
                "recipient@example.com",
                string.Empty,
                replyAll,
                "saved_to_drafts"));
        }
    }
}
