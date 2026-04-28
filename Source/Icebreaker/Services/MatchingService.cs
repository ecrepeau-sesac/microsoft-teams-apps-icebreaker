// <copyright file="MatchingService.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Icebreaker.Services
{
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Icebreaker.Helpers;
    using Icebreaker.Helpers.AdaptiveCards;
    using Icebreaker.Interfaces;
    using Microsoft.ApplicationInsights;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// Orchestrates team-member pairing and notifications.
    /// </summary>
    public class MatchingService : IMatchingService
    {
        // How many weeks of pairing history to consider when avoiding re-pairings.
        private const int RecentPairingWindowWeeks = 8;

        private readonly IBotDataProvider dataProvider;
        private readonly ConversationHelper conversationHelper;
        private readonly TelemetryClient telemetryClient;
        private readonly BotAdapter botAdapter;
        private readonly int maxPairUpsPerTeam;
        private readonly string botDisplayName;

        /// <summary>
        /// Initializes a new instance of the <see cref="MatchingService"/> class.
        /// </summary>
        public MatchingService(IBotDataProvider dataProvider, ConversationHelper conversationHelper, TelemetryClient telemetryClient, BotAdapter botAdapter)
        {
            this.dataProvider = dataProvider;
            this.conversationHelper = conversationHelper;
            this.telemetryClient = telemetryClient;
            this.botAdapter = botAdapter;
            this.maxPairUpsPerTeam = Convert.ToInt32(ConfigurationManager.AppSettings["MaxPairUpsPerTeam"]);
            this.botDisplayName = ConfigurationManager.AppSettings["BotDisplayName"];
        }

        /// <inheritdoc/>
        public async Task<int> MakePairsAndNotifyAsync()
        {
            this.telemetryClient.TrackTrace("Making pairups");

            var installedTeamsCount = 0;
            var pairsNotifiedCount = 0;
            var usersNotifiedCount = 0;
            var dbMembersCount = 0;

            try
            {
                var teams = await this.dataProvider.GetInstalledTeamsAsync();
                installedTeamsCount = teams.Count;
                this.telemetryClient.TrackTrace($"Generating pairs for {installedTeamsCount} teams");

                // Load full user info once — needed for frequency filtering and pairing history.
                var allUsersInfo = await this.dataProvider.GetAllUsersInfoAsync();
                dbMembersCount = allUsersInfo.Count;

                foreach (var team in teams)
                {
                    this.telemetryClient.TrackTrace($"Pairing members of team {team.Id}");

                    try
                    {
                        var teamName = await this.conversationHelper.GetTeamNameByIdAsync(this.botAdapter, team);
                        var eligibleUsers = await this.GetEligibleUsersAsync(allUsersInfo, team);

                        var pairs = this.MakePairs(eligibleUsers, allUsersInfo);

                        foreach (var pair in pairs.Take(this.maxPairUpsPerTeam))
                        {
                            usersNotifiedCount += await this.NotifyPairAsync(team, teamName, pair, default(CancellationToken));
                            pairsNotifiedCount++;

                            // Record this pairing in each user's history.
                            var id1 = this.GetChannelUserObjectId(pair.Item1);
                            var id2 = this.GetChannelUserObjectId(pair.Item2);
                            if (id1 != null && id2 != null)
                            {
                                await this.dataProvider.RecordPairUpAsync(team.TenantId, id1, id2, team.ServiceUrl);
                            }

                            // Handle trio — notify the third member too.
                            if (pair.Item3 != null)
                            {
                                var id3 = this.GetChannelUserObjectId(pair.Item3);
                                if (id3 != null)
                                {
                                    await this.dataProvider.RecordPairUpAsync(team.TenantId, id1, id3, team.ServiceUrl);
                                    await this.dataProvider.RecordPairUpAsync(team.TenantId, id2, id3, team.ServiceUrl);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        this.telemetryClient.TrackTrace($"Error pairing up team members: {ex.Message}", SeverityLevel.Warning);
                        this.telemetryClient.TrackException(ex);
                    }
                }
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackTrace($"Error making pairups: {ex.Message}", SeverityLevel.Warning);
                this.telemetryClient.TrackException(ex);
            }

            var properties = new Dictionary<string, string>
            {
                { "InstalledTeamsCount", installedTeamsCount.ToString() },
                { "PairsNotifiedCount", pairsNotifiedCount.ToString() },
                { "UsersNotifiedCount", usersNotifiedCount.ToString() },
                { "DBMembersCount", dbMembersCount.ToString() },
            };
            this.telemetryClient.TrackEvent("ProcessedPairups", properties);
            this.telemetryClient.TrackTrace($"Made {pairsNotifiedCount} pairups, {usersNotifiedCount} notifications sent");

            return pairsNotifiedCount;
        }

        /// <summary>
        /// Sends pairup notification cards to both (or all three) members of a pair/trio.
        /// Returns the number of users successfully notified.
        /// </summary>
        private async Task<int> NotifyPairAsync(TeamInstallInfo teamModel, string teamName, Tuple<ChannelAccount, ChannelAccount, ChannelAccount> pair, CancellationToken cancellationToken)
        {
            this.telemetryClient.TrackTrace($"Sending pairup notification to {pair.Item1.Id} and {pair.Item2.Id}");

            var person1 = JObject.FromObject(pair.Item1).ToObject<TeamsChannelAccount>();
            var person2 = JObject.FromObject(pair.Item2).ToObject<TeamsChannelAccount>();

            var card1 = PairUpNotificationAdaptiveCard.GetCard(teamName, person1, person2, this.botDisplayName);
            var card2 = PairUpNotificationAdaptiveCard.GetCard(teamName, person2, person1, this.botDisplayName);

            var tasks = new List<Task<bool>>
            {
                this.conversationHelper.NotifyUserAsync(this.botAdapter, teamModel.ServiceUrl, teamModel.TeamId, MessageFactory.Attachment(card1), person1, teamModel.TenantId, cancellationToken),
                this.conversationHelper.NotifyUserAsync(this.botAdapter, teamModel.ServiceUrl, teamModel.TeamId, MessageFactory.Attachment(card2), person2, teamModel.TenantId, cancellationToken),
            };

            // Trio support: notify the third person about both other members.
            if (pair.Item3 != null)
            {
                var person3 = JObject.FromObject(pair.Item3).ToObject<TeamsChannelAccount>();
                this.telemetryClient.TrackTrace($"Sending trio notification to {pair.Item3.Id}");

                // Person3 gets a card that shows Person1 as their match (they can reach out to either).
                var card3 = PairUpNotificationAdaptiveCard.GetCard(teamName, person3, person1, this.botDisplayName);
                tasks.Add(this.conversationHelper.NotifyUserAsync(this.botAdapter, teamModel.ServiceUrl, teamModel.TeamId, MessageFactory.Attachment(card3), person3, teamModel.TenantId, cancellationToken));
            }

            var results = await Task.WhenAll(tasks);
            return results.Count(wasNotified => wasNotified);
        }

        /// <summary>
        /// Returns opted-in users whose pairing frequency says they should be included this run.
        /// </summary>
        private async Task<List<ChannelAccount>> GetEligibleUsersAsync(Dictionary<string, UserInfo> allUsersInfo, TeamInstallInfo teamInfo)
        {
            var members = await this.conversationHelper.GetTeamMembers(this.botAdapter, teamInfo);
            this.telemetryClient.TrackTrace($"Found {members.Count} members in team {teamInfo.TeamId}");

            return members
                .Where(member => member != null)
                .Where(member =>
                {
                    var aadId = this.GetChannelUserObjectId(member);
                    if (aadId == null)
                    {
                        return true; // Unknown users default to opted-in.
                    }

                    if (!allUsersInfo.TryGetValue(aadId, out var info))
                    {
                        return true; // Never seen before → default opted-in.
                    }

                    return info.OptedIn;
                })
                .ToList();
        }

        /// <summary>
        /// Pairs users into groups of two, with one optional trio when the count is odd.
        /// Prefers pairings where neither person appears in the other's recent history.
        /// </summary>
        private List<Tuple<ChannelAccount, ChannelAccount, ChannelAccount>> MakePairs(
            List<ChannelAccount> users,
            Dictionary<string, UserInfo> allUsersInfo)
        {
            if (users.Count < 2)
            {
                this.telemetryClient.TrackTrace("Pairs could not be made because there are fewer than 2 eligible users");
                return new List<Tuple<ChannelAccount, ChannelAccount, ChannelAccount>>();
            }

            this.telemetryClient.TrackTrace($"Making pairs among {users.Count} users");

            // Shuffle first so tie-breaking in the sort is random.
            this.Shuffle(users);

            // Sort by how recently each user was paired — users with older/no history go first,
            // making it more likely they get a fresh match.
            var cutoff = DateTime.UtcNow.AddDays(-RecentPairingWindowWeeks * 7);
            users = users
                .OrderBy(u =>
                {
                    var id = this.GetChannelUserObjectId(u);
                    if (id == null || !allUsersInfo.TryGetValue(id, out var info) || info.RecentPairUps == null)
                    {
                        return DateTime.MinValue;
                    }

                    var recent = info.RecentPairUps
                        .Where(h => h.PairedOn >= cutoff)
                        .OrderByDescending(h => h.PairedOn)
                        .FirstOrDefault();
                    return recent?.PairedOn ?? DateTime.MinValue;
                })
                .ToList();

            var pairs = new List<Tuple<ChannelAccount, ChannelAccount, ChannelAccount>>();
            var used = new HashSet<int>();

            for (int i = 0; i < users.Count; i++)
            {
                if (used.Contains(i))
                {
                    continue;
                }

                var idI = this.GetChannelUserObjectId(users[i]);
                var historyI = this.GetRecentPairIds(idI, allUsersInfo, cutoff);

                // Find the best available partner — prefer someone not in i's recent history.
                int bestJ = -1;
                for (int j = i + 1; j < users.Count; j++)
                {
                    if (used.Contains(j))
                    {
                        continue;
                    }

                    var idJ = this.GetChannelUserObjectId(users[j]);
                    bool freshPair = !historyI.Contains(idJ);

                    if (freshPair || bestJ == -1)
                    {
                        bestJ = j;
                        if (freshPair)
                        {
                            break; // First fresh candidate is good enough.
                        }
                    }
                }

                if (bestJ == -1)
                {
                    continue; // Only one unpaired user remains — absorbed into a trio below.
                }

                used.Add(i);
                used.Add(bestJ);
                pairs.Add(new Tuple<ChannelAccount, ChannelAccount, ChannelAccount>(users[i], users[bestJ], null));
            }

            // If there is exactly one leftover user, absorb them into the last pair as a trio.
            var leftover = Enumerable.Range(0, users.Count).Where(k => !used.Contains(k)).ToList();
            if (leftover.Count == 1 && pairs.Count > 0)
            {
                var last = pairs[pairs.Count - 1];
                pairs[pairs.Count - 1] = new Tuple<ChannelAccount, ChannelAccount, ChannelAccount>(last.Item1, last.Item2, users[leftover[0]]);
                this.telemetryClient.TrackTrace("Odd-member group: formed a trio for the last match");
            }

            return pairs;
        }

        private HashSet<string> GetRecentPairIds(string userId, Dictionary<string, UserInfo> allUsersInfo, DateTime cutoff)
        {
            if (userId == null || !allUsersInfo.TryGetValue(userId, out var info) || info.RecentPairUps == null)
            {
                return new HashSet<string>();
            }

            return new HashSet<string>(
                info.RecentPairUps
                    .Where(h => h.PairedOn >= cutoff)
                    .Select(h => h.PairedWithUserId)
                    .Where(id => id != null));
        }

        private string GetChannelUserObjectId(ChannelAccount account)
        {
            return JObject.FromObject(account).ToObject<TeamsChannelAccount>()?.AadObjectId;
        }

        private void Shuffle<T>(IList<T> items)
        {
            var rand = new Random();
            for (int i = items.Count - 1; i > 0; i--)
            {
                int j = rand.Next(i + 1);
                T temp = items[i];
                items[i] = items[j];
                items[j] = temp;
            }
        }
    }
}
