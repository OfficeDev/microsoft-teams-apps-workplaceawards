// <copyright file="MentionActivityType.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Models
{
    /// <summary>
    /// Enum to specify the type of mention activity.
    /// </summary>
    public enum MentionActivityType
    {
        /// <summary>
        /// Set admin mention.
        /// </summary>
        SetAdmin = 0,

        /// <summary>
        /// Nomination mention.
        /// </summary>
        Nomination = 1,

        /// <summary>
        /// Winner mention.
        /// </summary>
        Winner = 2,
    }
}
