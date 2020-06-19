// <copyright file="IRewardCycleStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Providers
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;

    /// <summary>
    /// Interface for reward cycle storage provider.
    /// </summary>
    public interface IRewardCycleStorageProvider
    {
        /// <summary>
        /// This method is used to fetch reward cycle details for a given team Id.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <returns>Reward cycle for a given team Id.</returns>
        Task<RewardCycleEntity> GetCurrentRewardCycleAsync(string teamId);

        /// <summary>
        /// This method is used to fetch punished reward cycle details for a given team Id.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <returns>Reward cycle for a given team Id.</returns>
        Task<RewardCycleEntity> GetPublishedRewardCycleAsync(string teamId);

        /// <summary>
        /// Store or update reward cycle in table storage.
        /// </summary>
        /// <param name="rewardCycleEntity">Represents reward cycle entity used for storage and retrieval.</param>
        /// <returns><see cref="Task"/> that represents reward cycle entity is saved or updated.</returns>
        Task<RewardCycleEntity> StoreOrUpdateRewardCycleAsync(RewardCycleEntity rewardCycleEntity);

        /// <summary>
        /// This method is used to fetch current reward cycle detail for all teams.
        /// </summary>
        /// <returns><see cref="Task"/>Current reward cycle entities for all teams.</returns>
        Task<List<RewardCycleEntity>> GetCurrentRewardCycleForAllTeamsAsync();

        /// <summary>
        /// This method is used get active reward cycle details for all teams.
        /// </summary>
        /// <returns>Reward active cycle for all teams.</returns>
        Task<List<RewardCycleEntity>> GetActiveRewardCycleForAllTeamsAsync();
    }
}