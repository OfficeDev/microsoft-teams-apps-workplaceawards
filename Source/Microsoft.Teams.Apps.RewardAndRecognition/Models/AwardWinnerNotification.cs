// <copyright file="AwardWinnerNotification.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Models
{
    using Newtonsoft.Json;

    /// <summary>
    /// Awards winner entity.
    /// </summary>
    public class AwardWinnerNotification : PublishResult
    {
        /// <summary>
        /// Gets or sets award image URL.
        /// </summary>
        [JsonProperty("AwardLink")]
        public string AwardLink { get; set; }
    }
}
