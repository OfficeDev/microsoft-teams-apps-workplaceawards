// <copyright file="PublishResult.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Models
{
    using Newtonsoft.Json;

    /// <summary>
    /// Class contains details of awards winners.
    /// </summary>
    public class PublishResult : NominationEntity
    {
        /// <summary>
        /// Gets or sets award cycle.
        /// </summary>
        [JsonProperty("AwardCycle")]
        public string AwardCycle { get; set; }

        /// <summary>
        /// Gets or sets endorsement count for award nomination.
        /// </summary>
        [JsonProperty("EndorsementCount")]
        public int EndorsementCount { get; set; }
    }
}
