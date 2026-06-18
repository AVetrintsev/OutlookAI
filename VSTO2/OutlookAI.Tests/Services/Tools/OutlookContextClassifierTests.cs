using System;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public sealed class OutlookContextClassifierTests
    {
        private static readonly MailboxIdentity Me = new MailboxIdentity
        {
            DisplayName = "Ivan Petrov",
            SmtpAddress = "ivan@example.com",
            Aliases = new[] { "i.petrov@example.com" }
        };

        [Fact]
        public void MailDirection_FromMe_IsOutgoing()
        {
            Assert.Equal("outgoing", OutlookContextClassifier.MailDirection(
                Me,
                "Ivan <ivan@example.com>",
                new[] { "alice@example.com" },
                Array.Empty<string>()));
        }

        [Fact]
        public void MailDirection_ToMe_IsIncoming()
        {
            Assert.Equal("incoming", OutlookContextClassifier.MailDirection(
                Me,
                "Alice <alice@example.com>",
                new[] { "Ivan <i.petrov@example.com>" },
                Array.Empty<string>()));
        }

        [Fact]
        public void MailRole_CcMe_IsCcRecipient()
        {
            Assert.Equal("cc_recipient", OutlookContextClassifier.MailRole(
                Me,
                "Alice <alice@example.com>",
                Array.Empty<string>(),
                new[] { "Ivan <ivan@example.com>" }));
        }

        [Fact]
        public void MailDirection_NoIdentityMatch_IsUnknown()
        {
            Assert.Equal("unknown", OutlookContextClassifier.MailDirection(
                Me,
                "Alice <alice@example.com>",
                new[] { "bob@example.com" },
                Array.Empty<string>()));
        }

        [Fact]
        public void MeetingDirection_OrganizerMe_IsOutgoing()
        {
            Assert.Equal("outgoing", OutlookContextClassifier.MeetingDirection(
                Me,
                "Ivan <ivan@example.com>",
                new[] { "alice@example.com" },
                Array.Empty<string>()));
        }

        [Fact]
        public void MeetingRole_OptionalAttendeeMe_IsOptionalAttendee()
        {
            Assert.Equal("optional_attendee", OutlookContextClassifier.MeetingRole(
                Me,
                "Alice <alice@example.com>",
                Array.Empty<string>(),
                new[] { "Ivan <i.petrov@example.com>" }));
        }
    }
}
