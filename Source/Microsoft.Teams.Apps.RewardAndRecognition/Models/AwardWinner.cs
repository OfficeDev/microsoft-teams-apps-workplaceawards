// <copyright file="AwardWinner.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Models
{
    using System.Collections.Generic;
    using Newtonsoft.Json;

    /// <summary>
    /// Awards winner entity.
    /// </summary>
    public class AwardWinner
    {
        /// <summary>
        /// Gets or sets team id.
        /// </summary>
        [JsonProperty("TeamId")]
        public string TeamId { get; set; }

        /// <summary>
        /// Gets or sets winners.
        /// </summary>
        [JsonProperty("Winners")]
        public IEnumerable<AwardWinnerNotification> Winners { get; set; }
    }
}
