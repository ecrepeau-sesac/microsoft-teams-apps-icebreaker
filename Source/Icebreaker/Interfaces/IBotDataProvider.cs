// <copyright file="IBotDataProvider.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Icebreaker.Interfaces
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Icebreaker.Helpers;

    /// <summary>
    /// Data provider routines.
    /// </summary>
    public interface IBotDataProvider
    {
        /// <summary>
        /// Get the list of teams to which the app was installed.
        /// </summary>
        Task<IList<TeamInstallInfo>> GetInstalledTeamsAsync();

        /// <summary>
        /// Returns the team that the bot has been installed to.
        /// </summary>
        Task<TeamInstallInfo> GetInstalledTeamAsync(string teamId);

        /// <summary>
        /// Updates team installation status in store.
        /// </summary>
        Task UpdateTeamInstallStatusAsync(TeamInstallInfo team, bool installed);

        /// <summary>
        /// Get opt-in status for all known users (userId → optedIn).
        /// </summary>
        Task<Dictionary<string, bool>> GetAllUsersOptInStatusAsync();

        /// <summary>
        /// Get full <see cref="UserInfo"/> for all known users (userId → UserInfo).
        /// </summary>
        Task<Dictionary<string, UserInfo>> GetAllUsersInfoAsync();

        /// <summary>
        /// Get the stored information about a single user. Returns null if not found.
        /// </summary>
        Task<UserInfo> GetUserInfoAsync(string userId);

        /// <summary>
        /// Persist opt-in status (and service URL) for a user.
        /// </summary>
        Task SetUserInfoAsync(string tenantId, string userId, bool optedIn, string serviceUrl);

        /// <summary>
        /// Persist pairing-frequency preference for a user.
        /// </summary>
        Task SetUserFrequencyAsync(string tenantId, string userId, PairingFrequency frequency, string serviceUrl);

        /// <summary>
        /// Record a pairing between two users and persist updated history for both.
        /// </summary>
        Task RecordPairUpAsync(string tenantId, string user1Id, string user2Id, string serviceUrl);

        /// <summary>
        /// Record whether a user met their most-recent match.
        /// </summary>
        Task RecordMeetupFeedbackAsync(string tenantId, string userId, string pairedWithUserId, bool didMeet, string serviceUrl);
    }
}
