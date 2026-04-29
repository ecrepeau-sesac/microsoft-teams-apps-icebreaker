// <copyright file="MeetupFeedbackAdaptiveCard.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Icebreaker.Helpers.AdaptiveCards
{
    using System;
    using global::AdaptiveCards.Templating;
    using Icebreaker.Properties;
    using Microsoft.Bot.Schema;

    /// <summary>
    /// Follow-up card sent after a pairing to ask whether the users actually met.
    /// </summary>
    public class MeetupFeedbackAdaptiveCard : AdaptiveCardBase
    {
        private static readonly Lazy<AdaptiveCardTemplate> AdaptiveCardTemplate =
            new Lazy<AdaptiveCardTemplate>(() => CardTemplateHelper.GetAdaptiveCardTemplate(AdaptiveCardName.MeetupFeedback));

        /// <summary>
        /// Creates the meetup feedback card.
        /// </summary>
        /// <param name="matchName">Display name of the person the recipient was paired with.</param>
        /// <param name="pairedWithUserId">AAD object id of the paired-with user (embedded in action payload).</param>
        /// <returns>The feedback card attachment.</returns>
        public static Attachment GetCard(string matchName, string pairedWithUserId)
        {
            var cardData = new
            {
                feedbackCardTitle = string.Format(Resources.FeedbackCardTitle, matchName),
                feedbackCardBody = Resources.FeedbackCardBody,
                feedbackYesButtonText = Resources.FeedbackYesButtonText,
                feedbackNoButtonText = Resources.FeedbackNoButtonText,
                pairedWithUserId,
            };

            return GetCard(AdaptiveCardTemplate.Value, cardData);
        }
    }
}
