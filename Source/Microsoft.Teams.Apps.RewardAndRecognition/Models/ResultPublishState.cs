// <copyright file="ResultPublishState.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Models
{
    /// <summary>
    /// Enum to specify the reward publish state.
    /// </summary>
    public enum ResultPublishState
    {
        /// <summary>
        /// Unpublished award cycle.
        /// </summary>
        Unpublished = 0,

        /// <summary>
        /// Published award.
        /// </summary>
        Published = 1,
    }
}
