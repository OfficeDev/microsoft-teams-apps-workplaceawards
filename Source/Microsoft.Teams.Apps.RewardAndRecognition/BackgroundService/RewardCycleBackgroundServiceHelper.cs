// <copyright file="RewardCycleBackgroundServiceHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.BackgroundService
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;
    using Microsoft.Teams.Apps.RewardAndRecognition.Providers;

    /// <summary>
    /// Helper class to handle the reward cycle background service helper methods.
    /// </summary>
    public class RewardCycleBackgroundServiceHelper : IRewardCycleBackgroundServiceHelper
    {
        /// <summary>
        /// Helper for storing reward cycle details to azure table storage.
        /// </summary>
        private readonly IRewardCycleStorageProvider rewardCycleStorageProvider;

        /// <summary>
        /// Provider to store logs in Azure Application Insights.
        /// </summary>
        private readonly ILogger<RewardCycleBackgroundServiceHelper> logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="RewardCycleBackgroundServiceHelper"/> class.
        /// </summary>
        /// <param name="rewardCycleStorageProvider">Reward cycle storage provider.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        public RewardCycleBackgroundServiceHelper(
            IRewardCycleStorageProvider rewardCycleStorageProvider,
            ILogger<RewardCycleBackgroundServiceHelper> logger)
        {
            this.rewardCycleStorageProvider = rewardCycleStorageProvider;
            this.logger = logger;
        }

        /// <summary>
        /// This method is used to start/stop reward cycle.
        /// If reward cycle start date matches the current date this method will set the cycle sate to active.
        /// Reward cycle state will be set as inactive based on cycle range of recurrence state configured for the team.
        /// </summary>
        /// <returns><see cref="Task"/> that represents reward cycle entity is saved or updated.</returns>
        public async Task<bool> UpdateCycleStatusAsync()
        {
            var currentRewardCycles = await this.rewardCycleStorageProvider.GetCurrentRewardCycleForAllTeamsAsync();

            // update reward cycle state
            foreach (RewardCycleEntity currentCycle in currentRewardCycles)
            {
                try
                {
                    var newCycle = this.CheckAndUpdateRewardCycleState(currentCycle);

                    await this.rewardCycleStorageProvider.StoreOrUpdateRewardCycleAsync(newCycle);
                    this.logger.LogInformation($"Reward cycle set to {(RewardCycleState)newCycle.RewardCycleState} TeamId: {newCycle.TeamId}");
                }
#pragma warning disable CA1031 // Catching general exceptions to unblock updating reward cycle for next reward cycle iteration.
                catch (Exception ex)
#pragma warning restore CA1031 // Catching general exceptions to unblock updating reward cycle for next reward cycle iteration.
                {
                    this.logger.LogError(ex, $"Error occurred while updating reward cycle state for team: {currentCycle.TeamId}.");
                }
            }

            return true;
        }

        /// <summary>
        /// Update current reward cycle recurrence based on RecurrenceType:(RepeatIndefinitely / RepeatUntilEndDate / RepeatUntilOccurrenceCount).
        /// </summary>
        /// <param name="currentCycle">Current reward cycle for team</param>
        /// <returns>Returns updated reward cycle entity</returns>
        private RewardCycleEntity CheckAndUpdateRewardCycleState(RewardCycleEntity currentCycle)
        {
            DateTime currentUtcTime = DateTime.UtcNow;
            if (currentCycle.Recurrence == (int)RecurrenceType.SingleOccurrence)
            {
                // current date should be between start date and end date
                if (currentUtcTime >= currentCycle.RewardCycleStartDate.Date
                    && currentUtcTime <= currentCycle.RewardCycleEndDate.Date
                    && currentCycle.ResultPublished != (int)ResultPublishState.Published)
                {
                    currentCycle.RewardCycleState = (int)RewardCycleState.Active;
                }
                else
                {
                    currentCycle.RewardCycleState = (int)RewardCycleState.Inactive;
                }
            }
            else
            {
                var occurrenceType = (RecurrenceType)currentCycle.Recurrence;

                switch (occurrenceType)
                {
                    case RecurrenceType.RepeatIndefinitely:
                        if (currentUtcTime > currentCycle.RewardCycleEndDate.Date)
                        {
                            // set a new award cycle for same duration.
                            this.UpdateRewardCycleState(currentCycle);
                        }

                        break;
                    case RecurrenceType.RepeatUntilEndDate:
                        currentCycle.RangeOfOccurrenceEndDate = currentCycle.RangeOfOccurrenceEndDate?.Date.ToUniversalTime();
                        int cycleDurationInDays = (currentCycle.RewardCycleEndDate.Date - currentCycle.RewardCycleStartDate.Date).Days;
                        int? remainingDaysInOccurrenceEndDate = (currentCycle.RangeOfOccurrenceEndDate?.Date - currentUtcTime)?.Days;

                        if (currentUtcTime <= currentCycle.RewardCycleEndDate.Date
                            && currentUtcTime >= currentCycle.RewardCycleStartDate.Date
                            && currentCycle.ResultPublished != (int)ResultPublishState.Published)
                        {
                            currentCycle.RewardCycleState = (int)RewardCycleState.Active;
                        }
                        else if (currentUtcTime > currentCycle.RewardCycleEndDate.Date
                            && currentUtcTime <= currentCycle.RangeOfOccurrenceEndDate?.Date
                            && remainingDaysInOccurrenceEndDate > cycleDurationInDays)
                        {
                            // set a new award cycle for same duration till occurrence end date.
                            this.UpdateRewardCycleState(currentCycle);
                        }
                        else
                        {
                            currentCycle.RewardCycleState = (int)RewardCycleState.Inactive;
                        }

                        break;
                    case RecurrenceType.RepeatUntilOccurrenceCount:
                        if (currentCycle.NumberOfOccurrences > 0
                            && (currentUtcTime > currentCycle.RewardCycleEndDate.Date))
                        {
                            this.UpdateRewardCycleState(currentCycle);
                            currentCycle.NumberOfOccurrences -= 1;
                        }
                        else if (currentCycle.NumberOfOccurrences >= 0 &&
                            currentUtcTime <= currentCycle.RewardCycleEndDate.Date
                            && currentCycle.ResultPublished != (int)ResultPublishState.Published)
                        {
                            currentCycle.RewardCycleState = (int)RewardCycleState.Active;
                        }
                        else
                        {
                            currentCycle.RewardCycleState = (int)RewardCycleState.Inactive;
                        }

                        break;
                }
            }

            return currentCycle;
        }

        /// <summary>
        /// Update reward cycle entity properties based on recurrence settings.
        /// </summary>
        /// <param name="currentCycle">Current reward cycle for team.</param>
        /// <returns>Returns new reward cycle entity.</returns>
        private RewardCycleEntity UpdateRewardCycleState(RewardCycleEntity currentCycle)
        {
            var guidValue = Guid.NewGuid().ToString();
            int cycleDurationInDays = (currentCycle.RewardCycleEndDate.Date - currentCycle.RewardCycleStartDate.Date).Days;

            currentCycle.CreatedOn = DateTime.UtcNow;
            currentCycle.CycleId = guidValue;
            currentCycle.ResultPublished = (int)ResultPublishState.Unpublished;
            currentCycle.RewardCycleEndDate = DateTime.UtcNow.AddDays(cycleDurationInDays);
            currentCycle.RewardCycleStartDate = DateTime.UtcNow;
            currentCycle.RewardCycleState = (int)RewardCycleState.Active;

            return currentCycle;
        }
    }
}
