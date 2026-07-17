using System.Text.Json;
using System.Text.Json.Serialization;
using OutlookComMcp.Models;
using OutlookComMcp.Outlook;
using OutlookComMcp.Tools;
using Xunit;

namespace OutlookComMcp.Tests;

public sealed class OutlookToolsTests
{
    [Fact]
    public void MailSummarySerializesNullBodyPreview()
    {
        MailSummary summary = new(
            "email-id",
            "store-id",
            "Subject",
            "Sender",
            "sender@example.com",
            DateTimeOffset.Parse("2026-07-17T09:00:00+09:00"),
            false,
            false,
            null);
        JsonSerializerOptions options = new()
        {
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        };

        string json = JsonSerializer.Serialize(summary, options);

        Assert.Contains("\"bodyPreview\":null", json, StringComparison.Ordinal);
    }

    [Fact]
    public async Task SearchEmailsPassesValidatedArgumentsToClient()
    {
        FakeOutlookClient client = new();
        OutlookTools tools = new(client);

        await tools.SearchEmailsAsync("inbox", "status", 14, 10, true, TestContext.Current.CancellationToken);

        Assert.Equal(("inbox", "status", 14, 10, true), client.SearchArguments);
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
                TestContext.Current.CancellationToken));
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
        public (string Folder, string? Query, int DaysBack, int MaxResults, bool IncludeBodyPreview)? SearchArguments { get; private set; }

        public (string EmailId, string StoreId, string Body, bool ReplyAll)? DraftArguments { get; private set; }

        public Task<IReadOnlyList<MailSummary>> SearchEmailsAsync(
            string folder,
            string? query,
            int daysBack,
            int maxResults,
            bool includeBodyPreview,
            CancellationToken cancellationToken)
        {
            SearchArguments = (folder, query, daysBack, maxResults, includeBodyPreview);
            return Task.FromResult<IReadOnlyList<MailSummary>>([]);
        }

        public Task<MailDetail> GetEmailAsync(
            string emailId,
            string storeId,
            int maxBodyCharacters,
            CancellationToken cancellationToken) => throw new NotImplementedException();

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
