// <copyright file="SendFeedbackController.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Icebreaker.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Web.Hosting;
    using System.Web.Http;
    using Icebreaker.Helpers;
    using Icebreaker.Helpers.AdaptiveCards;
    using Icebreaker.Interfaces;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// API endpoint that proactively sends post-meetup feedback cards to recently paired users.
    /// Call this from a Logic App or Azure Function N days after a pairing run.
    /// </summary>
    [RoutePrefix("api/sendfeedback")]
    public class SendFeedbackController : ApiController
    {
        private const string KeyHeaderName = "X-Key";

        // Only send feedback requests for pairings newer than this many days.
        private const int FeedbackWindowDays = 7;

        private readonly IBotDataProvider dataProvider;
        private readonly ConversationHelper conversationHelper;
        private readonly BotAdapter botAdapter;
        private readonly MicrosoftAppCredentials botCredentials;
        private readonly ISecretsHelper secretsHelper;

        /// <summary>
        /// Initializes a new instance of the <see cref="SendFeedbackController"/> class.
        /// </summary>
        public SendFeedbackController(
            IBotDataProvider dataProvider,
            ConversationHelper conversationHelper,
            BotAdapter botAdapter,
            MicrosoftAppCredentials botCredentials,
            ISecretsHelper secretsHelper)
        {
            this.dataProvider = dataProvider;
            this.conversationHelper = conversationHelper;
            this.botAdapter = botAdapter;
            this.botCredentials = botCredentials;
            this.secretsHelper = secretsHelper;
        }

        /// <summary>
        /// Triggers feedback cards for all users who have an unacknowledged recent pairing.
        /// </summary>
        [HttpGet]
        [Route("")]
        public async Task<IHttpActionResult> SendAsync()
        {
            IEnumerable<string> keys;
            if (!this.Request.Headers.TryGetValues(KeyHeaderName, out keys) ||
                !keys.Any() ||
                !object.Equals(keys.First(), this.secretsHelper.Key))
            {
                return this.Unauthorized();
            }

            await this.botCredentials.GetTokenAsync();

            HostingEnvironment.QueueBackgroundWorkItem(ct => this.SendFeedbackCardsAsync(ct));

            return this.StatusCode(HttpStatusCode.OK);
        }

        private async Task SendFeedbackCardsAsync(CancellationToken cancellationToken)
        {
            var cutoff = DateTime.UtcNow.AddDays(-FeedbackWindowDays);
            var allUsersInfo = await this.dataProvider.GetAllUsersInfoAsync();
            var installedTeams = await this.dataProvider.GetInstalledTeamsAsync();

            // Build a quick lookup of teamId → TeamInstallInfo for service URL resolution.
            var teamLookup = installedTeams.ToDictionary(t => t.TeamId, t => t);

            foreach (var kvp in allUsersInfo)
            {
                var userInfo = kvp.Value;
                if (userInfo.RecentPairUps == null)
                {
                    continue;
                }

                // Find pairings that are recent, have no feedback yet, and fall within the window.
                var pendingFeedback = userInfo.RecentPairUps
                    .Where(p => p.DidMeet == null && p.PairedOn >= cutoff)
                    .OrderByDescending(p => p.PairedOn)
                    .FirstOrDefault();

                if (pendingFeedback == null)
                {
                    continue;
                }

                // Look up the paired-with user's display name.
                if (!allUsersInfo.TryGetValue(pendingFeedback.PairedWithUserId, out var pairedUserInfo))
                {
                    continue;
                }

                // Find a team that both users are (or were) part of.
                var team = installedTeams.FirstOrDefault(t => t.TenantId == userInfo.TenantId);
                if (team == null)
                {
                    continue;
                }

                try
                {
                    // We use just the pairedWithUserId as the match name placeholder — in a full
                    // implementation you would resolve the display name via Graph or cache.
                    var card = MeetupFeedbackAdaptiveCard.GetCard(pendingFeedback.PairedWithUserId, pendingFeedback.PairedWithUserId);

                    var userChannelAccount = new TeamsChannelAccount { AadObjectId = userInfo.UserId };
                    await this.conversationHelper.NotifyUserAsync(
                        this.botAdapter,
                        userInfo.ServiceUrl ?? team.ServiceUrl,
                        team.TeamId,
                        MessageFactory.Attachment(card),
                        userChannelAccount,
                        userInfo.TenantId,
                        cancellationToken);
                }
                catch (Exception)
                {
                    // Best-effort — continue with other users.
                }
            }
        }
    }
}
