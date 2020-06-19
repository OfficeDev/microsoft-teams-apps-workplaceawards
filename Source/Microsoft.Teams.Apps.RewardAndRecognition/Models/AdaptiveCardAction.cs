// <copyright file="AdaptiveCardAction.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Models
{
    using Microsoft.Bot.Schema;
    using Newtonsoft.Json;

    /// <summary>
    /// Adaptive card action model class.
    /// </summary>
    public class AdaptiveCardAction
    {
        /// <summary>
        /// Gets or sets Ms Teams card action type.
        /// </summary>
        [JsonProperty("msteams")]
        public CardAction MsteamsCardAction { get; set; }

        /// <summary>
        /// Gets or sets name of award.
        /// </summary>
        [JsonProperty("AwardName")]
        public string AwardName { get; set; }

        /// <summary>
        /// Gets or sets name of award id.
        /// </summary>
        [JsonProperty("AwardId")]
        public string AwardId { get; set; }

        /// <summary>
        /// Gets or sets nominee name.
        /// </summary>
        [JsonProperty("NomineeNames")]
        public string NomineeNames { get; set; }

        /// <summary>
        /// Gets or sets User principal name of nominee.
        /// </summary>
        [JsonProperty("NomineeUserPrincipalNames")]
        public string NomineeUserPrincipalNames { get; set; }

        /// <summary>
        /// Gets or sets AAD object id of nominee.
        /// </summary>
        [JsonProperty("NomineeObjectIds")]
        public string NomineeObjectIds { get; set; }

        /// <summary>
        /// Gets or sets reward cycle identifier.
        /// </summary>
        [JsonProperty("RewardCycleId")]
        public string RewardCycleId { get; set; }

        /// <summary>
        /// Gets or sets commands from which task module is invoked.
        /// </summary>
        [JsonProperty("command")]
        public string Command { get; set; }

        /// <summary>
        /// Gets or sets group name.
        /// </summary>
        [JsonProperty("GroupName")]
        public string GroupName { get; set; }
    }
}
