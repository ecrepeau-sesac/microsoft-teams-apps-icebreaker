// <copyright file="IcebreakerBotDataProvider.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Icebreaker.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using System.Configuration;
    using Icebreaker.Interfaces;
    using Microsoft.ApplicationInsights;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.Azure.Documents;
    using Microsoft.Azure.Documents.Client;
    using Microsoft.Azure.Documents.Linq;

    /// <summary>
    /// CosmosDB-backed data provider.
    /// </summary>
    public class IcebreakerBotDataProvider : IBotDataProvider
    {
        private const int DefaultRequestThroughput = 400;
        private const int MaxPairUpHistoryEntries = 10;

        private readonly TelemetryClient telemetryClient;
        private readonly Lazy<Task> initializeTask;
        private readonly ISecretsHelper secretsHelper;
        private DocumentClient documentClient;
        private Database database;
        private DocumentCollection teamsCollection;
        private DocumentCollection usersCollection;

        /// <summary>
        /// Initializes a new instance of the <see cref="IcebreakerBotDataProvider"/> class.
        /// </summary>
        public IcebreakerBotDataProvider(TelemetryClient telemetryClient, ISecretsHelper secretsHelper)
        {
            this.telemetryClient = telemetryClient;
            this.secretsHelper = secretsHelper;
            this.initializeTask = new Lazy<Task>(() => this.InitializeAsync());
        }

        /// <inheritdoc/>
        public async Task UpdateTeamInstallStatusAsync(TeamInstallInfo team, bool installed)
        {
            await this.EnsureInitializedAsync();

            if (installed)
            {
                await this.documentClient.UpsertDocumentAsync(this.teamsCollection.SelfLink, team);
            }
            else
            {
                var documentUri = UriFactory.CreateDocumentUri(this.database.Id, this.teamsCollection.Id, team.Id);
                await this.documentClient.DeleteDocumentAsync(documentUri, new RequestOptions { PartitionKey = new PartitionKey(team.Id) });
            }
        }

        /// <inheritdoc/>
        public async Task<IList<TeamInstallInfo>> GetInstalledTeamsAsync()
        {
            await this.EnsureInitializedAsync();

            var installedTeams = new List<TeamInstallInfo>();

            try
            {
                using (var lookupQuery = this.documentClient
                    .CreateDocumentQuery<TeamInstallInfo>(this.teamsCollection.SelfLink, new FeedOptions { EnableCrossPartitionQuery = true })
                    .AsDocumentQuery())
                {
                    while (lookupQuery.HasMoreResults)
                    {
                        var response = await lookupQuery.ExecuteNextAsync<TeamInstallInfo>();
                        installedTeams.AddRange(response);
                    }
                }
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex.InnerException);
            }

            return installedTeams;
        }

        /// <inheritdoc/>
        public async Task<TeamInstallInfo> GetInstalledTeamAsync(string teamId)
        {
            await this.EnsureInitializedAsync();

            try
            {
                var documentUri = UriFactory.CreateDocumentUri(this.database.Id, this.teamsCollection.Id, teamId);
                return await this.documentClient.ReadDocumentAsync<TeamInstallInfo>(documentUri, new RequestOptions { PartitionKey = new PartitionKey(teamId) });
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex.InnerException);
                return null;
            }
        }

        /// <inheritdoc/>
        public async Task<UserInfo> GetUserInfoAsync(string userId)
        {
            await this.EnsureInitializedAsync();

            try
            {
                var documentUri = UriFactory.CreateDocumentUri(this.database.Id, this.usersCollection.Id, userId);
                return await this.documentClient.ReadDocumentAsync<UserInfo>(documentUri, new RequestOptions { PartitionKey = new PartitionKey(userId) });
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex.InnerException);
                return null;
            }
        }

        /// <inheritdoc/>
        public async Task<Dictionary<string, bool>> GetAllUsersOptInStatusAsync()
        {
            await this.EnsureInitializedAsync();

            try
            {
                var collectionLink = UriFactory.CreateDocumentCollectionUri(this.database.Id, this.usersCollection.Id);
                var query = this.documentClient.CreateDocumentQuery<UserInfo>(
                        collectionLink,
                        new FeedOptions { EnableCrossPartitionQuery = true, MaxItemCount = -1, MaxDegreeOfParallelism = -1 })
                    .Select(u => new UserInfo { Id = u.Id, OptedIn = u.OptedIn })
                    .AsDocumentQuery();

                var result = new Dictionary<string, bool>();
                while (query.HasMoreResults)
                {
                    var batch = await query.ExecuteNextAsync<UserInfo>();
                    foreach (var userInfo in batch)
                    {
                        result[userInfo.Id] = userInfo.OptedIn;
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex.InnerException);
                return new Dictionary<string, bool>();
            }
        }

        /// <inheritdoc/>
        public async Task<Dictionary<string, UserInfo>> GetAllUsersInfoAsync()
        {
            await this.EnsureInitializedAsync();

            try
            {
                var collectionLink = UriFactory.CreateDocumentCollectionUri(this.database.Id, this.usersCollection.Id);
                var query = this.documentClient.CreateDocumentQuery<UserInfo>(
                        collectionLink,
                        new FeedOptions { EnableCrossPartitionQuery = true, MaxItemCount = -1, MaxDegreeOfParallelism = -1 })
                    .AsDocumentQuery();

                var result = new Dictionary<string, UserInfo>();
                while (query.HasMoreResults)
                {
                    var batch = await query.ExecuteNextAsync<UserInfo>();
                    foreach (var userInfo in batch)
                    {
                        result[userInfo.Id] = userInfo;
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex.InnerException);
                return new Dictionary<string, UserInfo>();
            }
        }

        /// <inheritdoc/>
        public async Task SetUserInfoAsync(string tenantId, string userId, bool optedIn, string serviceUrl)
        {
            await this.EnsureInitializedAsync();

            var existing = await this.GetUserInfoAsync(userId);
            var userInfo = existing ?? new UserInfo();
            userInfo.TenantId = tenantId;
            userInfo.UserId = userId;
            userInfo.OptedIn = optedIn;
            userInfo.ServiceUrl = serviceUrl;

            await this.documentClient.UpsertDocumentAsync(this.usersCollection.SelfLink, userInfo);
        }

        /// <inheritdoc/>
        public async Task SetUserFrequencyAsync(string tenantId, string userId, PairingFrequency frequency, string serviceUrl)
        {
            await this.EnsureInitializedAsync();

            var existing = await this.GetUserInfoAsync(userId);
            var userInfo = existing ?? new UserInfo();
            userInfo.TenantId = tenantId;
            userInfo.UserId = userId;
            userInfo.ServiceUrl = serviceUrl;
            userInfo.PairingFrequency = frequency;

            await this.documentClient.UpsertDocumentAsync(this.usersCollection.SelfLink, userInfo);
        }

        /// <inheritdoc/>
        public async Task RecordPairUpAsync(string tenantId, string user1Id, string user2Id, string serviceUrl)
        {
            await this.EnsureInitializedAsync();

            await Task.WhenAll(
                this.AppendPairUpHistoryAsync(tenantId, user1Id, user2Id, serviceUrl),
                this.AppendPairUpHistoryAsync(tenantId, user2Id, user1Id, serviceUrl));
        }

        /// <inheritdoc/>
        public async Task RecordMeetupFeedbackAsync(string tenantId, string userId, string pairedWithUserId, bool didMeet, string serviceUrl)
        {
            await this.EnsureInitializedAsync();

            var userInfo = await this.GetUserInfoAsync(userId);
            if (userInfo == null)
            {
                return;
            }

            var entry = userInfo.RecentPairUps?.LastOrDefault(p => p.PairedWithUserId == pairedWithUserId && p.DidMeet == null);
            if (entry != null)
            {
                entry.DidMeet = didMeet;
                await this.documentClient.UpsertDocumentAsync(this.usersCollection.SelfLink, userInfo);
            }
        }

        private async Task AppendPairUpHistoryAsync(string tenantId, string userId, string pairedWithUserId, string serviceUrl)
        {
            var userInfo = await this.GetUserInfoAsync(userId) ?? new UserInfo
            {
                TenantId = tenantId,
                UserId = userId,
                ServiceUrl = serviceUrl,
                OptedIn = true,
            };

            if (userInfo.RecentPairUps == null)
            {
                userInfo.RecentPairUps = new System.Collections.Generic.List<PairUpHistory>();
            }

            userInfo.RecentPairUps.Add(new PairUpHistory
            {
                PairedWithUserId = pairedWithUserId,
                PairedOn = DateTime.UtcNow,
            });

            // Keep only the most recent N entries
            if (userInfo.RecentPairUps.Count > MaxPairUpHistoryEntries)
            {
                userInfo.RecentPairUps = userInfo.RecentPairUps
                    .OrderByDescending(h => h.PairedOn)
                    .Take(MaxPairUpHistoryEntries)
                    .ToList();
            }

            await this.documentClient.UpsertDocumentAsync(this.usersCollection.SelfLink, userInfo);
        }

        private async Task InitializeAsync()
        {
            this.telemetryClient.TrackTrace("Initializing data store");

            var endpointUrl = ConfigurationManager.AppSettings["CosmosDBEndpointUrl"];
            var databaseName = ConfigurationManager.AppSettings["CosmosDBDatabaseName"];
            var teamsCollectionName = ConfigurationManager.AppSettings["CosmosCollectionTeams"];
            var usersCollectionName = ConfigurationManager.AppSettings["CosmosCollectionUsers"];

            this.documentClient = new DocumentClient(new Uri(endpointUrl), this.secretsHelper.CosmosDBKey);

            var requestOptions = new RequestOptions { OfferThroughput = DefaultRequestThroughput };
            bool useSharedOffer = true;

            try
            {
                this.database = await this.documentClient.CreateDatabaseIfNotExistsAsync(new Database { Id = databaseName }, requestOptions);
            }
            catch (DocumentClientException ex)
            {
                if (ex.Error?.Message?.Contains("SharedOffer is Disabled") ?? false)
                {
                    this.telemetryClient.TrackTrace("Database shared offer is disabled, provisioning throughput at container level", SeverityLevel.Information);
                    useSharedOffer = false;
                    this.database = await this.documentClient.CreateDatabaseIfNotExistsAsync(new Database { Id = databaseName });
                }
                else
                {
                    throw;
                }
            }

            var teamsCollectionDefinition = new DocumentCollection { Id = teamsCollectionName };
            teamsCollectionDefinition.PartitionKey.Paths.Add("/id");
            this.teamsCollection = await this.documentClient.CreateDocumentCollectionIfNotExistsAsync(this.database.SelfLink, teamsCollectionDefinition, useSharedOffer ? null : requestOptions);

            var usersCollectionDefinition = new DocumentCollection { Id = usersCollectionName };
            usersCollectionDefinition.PartitionKey.Paths.Add("/id");
            this.usersCollection = await this.documentClient.CreateDocumentCollectionIfNotExistsAsync(this.database.SelfLink, usersCollectionDefinition, useSharedOffer ? null : requestOptions);

            this.telemetryClient.TrackTrace("Data store initialized");
        }

        private async Task EnsureInitializedAsync()
        {
            await this.initializeTask.Value;
        }
    }
}
