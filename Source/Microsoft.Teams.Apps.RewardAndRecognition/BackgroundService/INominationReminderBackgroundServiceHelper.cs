// <copyright file="INominationReminderBackgroundServiceHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.BackgroundService
{
    using System.Threading.Tasks;

    /// <summary>
    /// Interface to provide helper methods implementation for nomination reminder
    /// notifications used in background service.
    /// </summary>
    public interface INominationReminderBackgroundServiceHelper
    {
        /// <summary>
        /// This method is used to send nomination reminder notification to Teams channel.
        /// </summary>
        /// <returns>A <see cref="Task"/>Returns true if reward cycle set successfully, else false.</returns>
        Task<bool> SendNominationReminderNotificationAsync();
    }
}
