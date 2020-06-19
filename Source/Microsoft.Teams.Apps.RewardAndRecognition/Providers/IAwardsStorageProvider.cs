// <copyright file="IAwardsStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Providers
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;

    /// <summary>
    /// Awards storage provider interface.
    /// </summary>
    public interface IAwardsStorageProvider
    {
        /// <summary>
        /// This method is used to fetch all the awards for a given team Id.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <returns>Returns all the awards for a given team Id.</returns>
        Task<IEnumerable<AwardEntity>> GetAwardsAsync(string teamId);

        /// <summary>
        /// Store or update awards in table storage.
        /// </summary>
        /// <param name="awardEntity">Represents award entity used for storage and retrieval.</param>
        /// <returns><see cref="Task"/>Returns award entity is saved or updated.</returns>
        Task<AwardEntity> StoreOrUpdateAwardAsync(AwardEntity awardEntity);

        /// <summary>
        /// Delete award details data in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="teamId">Holds team Id.</param>
        /// <param name="awardIds">Holds award Id data.</param>
        /// <returns>A task that represents award entity data is saved or updated.</returns>
        Task<bool> DeleteAwardsAsync(string teamId, IEnumerable<string> awardIds);

        /// <summary>
        /// This method is used to fetch award details for a given team Id and awardId.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <param name="awardId">Award Id.</param>
        /// <returns>Award details.</returns>
        Task<AwardEntity> GetAwardDetailsAsync(string teamId, string awardId);
    }
}