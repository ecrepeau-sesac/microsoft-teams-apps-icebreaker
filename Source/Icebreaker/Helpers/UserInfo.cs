// <copyright file="UserInfo.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Icebreaker.Helpers
{
    using System.Collections.Generic;
    using Microsoft.Azure.Documents;
    using Newtonsoft.Json;

    /// <summary>
    /// Represents a user and their Icebreaker preferences.
    /// </summary>
    public class UserInfo : Document
    {
        /// <summary>
        /// Gets or sets the user's AAD object id.
        /// This is also the <see cref="Resource.Id"/>.
        /// </summary>
        [JsonIgnore]
        public string UserId
        {
            get { return this.Id; }
            set { this.Id = value; }
        }

        /// <summary>
        /// Gets or sets the tenant id.
        /// </summary>
        [JsonProperty("tenantId")]
        public string TenantId { get; set; }

        /// <summary>
        /// Gets or sets the service URL.
        /// </summary>
        [JsonProperty("serviceUrl")]
        public string ServiceUrl { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the user is opted in to pairups.
        /// </summary>
        [JsonProperty("optedIn")]
        public bool OptedIn { get; set; }

        /// <summary>
        /// Gets or sets how often the user wants to be paired.
        /// Defaults to <see cref="PairingFrequency.Weekly"/>.
        /// </summary>
        [JsonProperty("pairingFrequency")]
        public PairingFrequency PairingFrequency { get; set; } = PairingFrequency.Weekly;

        /// <summary>
        /// Gets or sets the pairing run counter used to implement non-weekly frequencies.
        /// Incremented each global pairing run; the user is included only when
        /// (RunCounter % (int)PairingFrequency == 0).
        /// </summary>
        [JsonProperty("lastPairingRun")]
        public int LastPairingRun { get; set; }

        /// <summary>
        /// Gets or sets the history of recent pairups, capped at the last 10 entries.
        /// </summary>
        [JsonProperty("recentPairups")]
        public List<PairUpHistory> RecentPairUps { get; set; } = new List<PairUpHistory>();
    }
}
