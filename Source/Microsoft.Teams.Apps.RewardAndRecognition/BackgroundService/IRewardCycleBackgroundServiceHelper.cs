// <copyright file="IRewardCycleBackgroundServiceHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.BackgroundService
{
    using System.Threading.Tasks;

    /// <summary>
    /// Interface to provide helper method implementations for updating reward cycle
    /// by background service.
    /// </summary>
    public interface IRewardCycleBackgroundServiceHelper
    {
        /// <summary>
        /// This method is used to start reward cycle if the start date matches the current date and stops the reward cycle if the end date matches the current date.
        /// Update current reward cycle recurrence based on RecurrenceType:(RepeatIndefinitely / RepeatUntilEndDate / RepeatUntilOccurrenceCount).
        /// </summary>
        /// <returns><see cref="Task"/> that represents reward cycle entity is saved or updated.</returns>
        Task<bool> UpdateCycleStatusAsync();
    }
}
