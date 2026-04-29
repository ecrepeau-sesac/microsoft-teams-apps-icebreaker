// <copyright file="MatchingActions.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Icebreaker.Helpers
{
    /// <summary>
    /// Message-back text constants for all bot actions.
    /// </summary>
    public static class MatchingActions
    {
        /// <summary>Opt in to pair matching.</summary>
        public const string OptIn = "optin";

        /// <summary>Opt out of pair matching.</summary>
        public const string OptOut = "optout";

        /// <summary>User confirms they met their match.</summary>
        public const string FeedbackYes = "feedback_yes";

        /// <summary>User indicates they have not met their match yet.</summary>
        public const string FeedbackNo = "feedback_no";

        /// <summary>Set pairing frequency to weekly.</summary>
        public const string FrequencyWeekly = "frequency_weekly";

        /// <summary>Set pairing frequency to every two weeks.</summary>
        public const string FrequencyBiweekly = "frequency_biweekly";

        /// <summary>Set pairing frequency to monthly.</summary>
        public const string FrequencyMonthly = "frequency_monthly";
    }
}
