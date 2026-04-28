// <copyright file="PairingFrequency.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Icebreaker.Helpers
{
    /// <summary>
    /// How often a user wants to be paired.
    /// The integer value equals the number of weeks between pairings.
    /// </summary>
    public enum PairingFrequency
    {
        /// <summary>Paired every week (default).</summary>
        Weekly = 1,

        /// <summary>Paired every two weeks.</summary>
        Biweekly = 2,

        /// <summary>Paired approximately once a month (every four weeks).</summary>
        Monthly = 4,
    }
}
