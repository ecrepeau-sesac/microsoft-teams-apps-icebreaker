// <copyright file="PairUpNotificationAdaptiveCard.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Icebreaker.Helpers.AdaptiveCards
{
    using System;
    using System.Globalization;
    using global::AdaptiveCards;
    using global::AdaptiveCards.Templating;
    using Icebreaker.Properties;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;

    /// <summary>
    /// Builder class for the pairup notification card.
    /// </summary>
    public class PairUpNotificationAdaptiveCard : AdaptiveCardBase
    {
        private const string ExternallyAuthenticatedUpnMarker = "#ext#";

        private static readonly Lazy<AdaptiveCardTemplate> AdaptiveCardTemplate =
            new Lazy<AdaptiveCardTemplate>(() => CardTemplateHelper.GetAdaptiveCardTemplate(AdaptiveCardName.PairUpNotification));

        /// <summary>
        /// Creates the pairup notification card.
        /// </summary>
        /// <param name="teamName">The team name.</param>
        /// <param name="sender">The user receiving this card (they see the recipient as their match).</param>
        /// <param name="recipient">The user who is the match.</param>
        /// <param name="botDisplayName">The bot display name.</param>
        /// <param name="currentFrequency">The sender's current pairing frequency (used to highlight active button).</param>
        /// <returns>Pairup notification card attachment.</returns>
        public static Attachment GetCard(string teamName, TeamsChannelAccount sender, TeamsChannelAccount recipient, string botDisplayName, PairingFrequency currentFrequency = PairingFrequency.Weekly)
        {
            var textAlignment = CultureInfo.CurrentCulture.TextInfo.IsRightToLeft ? AdaptiveHorizontalAlignment.Right.ToString() : AdaptiveHorizontalAlignment.Left.ToString();

            var senderGivenName = string.IsNullOrEmpty(sender.GivenName) ? sender.Name : sender.GivenName;
            var recipientGivenName = string.IsNullOrEmpty(recipient.GivenName) ? recipient.Name : recipient.GivenName;

            var recipientUpn = !IsGuestUser(recipient) ? recipient.UserPrincipalName : recipient.Email;

            var meetingTitle = string.Format(Resources.MeetupTitle, senderGivenName, recipientGivenName);
            var meetingContent = string.Format(Resources.MeetupContent, botDisplayName);
            var meetingLink = "https://teams.microsoft.com/l/meeting/new?subject=" + Uri.EscapeDataString(meetingTitle) + "&attendees=" + recipientUpn + "&content=" + Uri.EscapeDataString(meetingContent);

            var cardData = new
            {
                matchUpCardTitleContent = Resources.MatchUpCardTitleContent,
                matchUpCardMatchedText = string.Format(Resources.MatchUpCardMatchedText, recipient.Name),
                matchUpCardContentPart1 = string.Format(Resources.MatchUpCardContentPart1, botDisplayName, teamName, recipient.Name),
                matchUpCardContentPart2 = Resources.MatchUpCardContentPart2,
                chatWithMatchButtonText = string.Format(Resources.ChatWithMatchButtonText, recipientGivenName),
                chatWithMessageGreeting = Uri.EscapeDataString(Resources.ChatWithMessageGreeting),
                pauseMatchesButtonText = Resources.PausePairingsButtonText,
                proposeMeetupButtonText = Resources.ProposeMeetupButtonText,
                personUpn = recipientUpn,
                meetingLink,
                textAlignment,
                changeFrequencyLabel = Resources.ChangeFrequencyButtonText,
                frequencyWeeklyButtonText = Resources.FrequencyWeeklyButtonText,
                frequencyBiweeklyButtonText = Resources.FrequencyBiweeklyButtonText,
                frequencyMonthlyButtonText = Resources.FrequencyMonthlyButtonText,
                frequencyWeeklyStyle = currentFrequency == PairingFrequency.Weekly ? "positive" : "default",
                frequencyBiweeklyStyle = currentFrequency == PairingFrequency.Biweekly ? "positive" : "default",
                frequencyMonthlyStyle = currentFrequency == PairingFrequency.Monthly ? "positive" : "default",
            };

            return GetCard(AdaptiveCardTemplate.Value, cardData);
        }

        private static bool IsGuestUser(TeamsChannelAccount account)
        {
            return account.UserPrincipalName.IndexOf(ExternallyAuthenticatedUpnMarker, StringComparison.InvariantCultureIgnoreCase) >= 0;
        }
    }
}
