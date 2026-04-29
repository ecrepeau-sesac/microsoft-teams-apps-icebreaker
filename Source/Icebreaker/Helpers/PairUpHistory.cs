// <copyright file="PairUpHistory.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Icebreaker.Helpers
{
    using System;
    using Newtonsoft.Json;

    /// <summary>
    /// Records a single historical pairing between the owning user and another user.
    /// </summary>
    public class PairUpHistory
    {
        /// <summary>
        /// Gets or sets the AAD object id of the user this person was paired with.
        /// </summary>
        [JsonProperty("pairedWithUserId")]
        public string PairedWithUserId { get; set; }

        /// <summary>
        /// Gets or sets the UTC timestamp when this pairing was made.
        /// </summary>
        [JsonProperty("pairedOn")]
        public DateTime PairedOn { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the users confirmed they actually met.
        /// Null means no feedback received yet.
        /// </summary>
        [JsonProperty("didMeet")]
        public bool? DidMeet { get; set; }
    }
}
