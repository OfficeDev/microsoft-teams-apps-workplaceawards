// <copyright file="IEndorsementsStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Providers
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;

    /// <summary>
    /// Interface for endorsement storage provider.
    /// </summary>
    public interface IEndorsementsStorageProvider
    {
        /// <summary>
        /// Store or update endorsement in table storage.
        /// </summary>
        /// <param name="endorseEntity">Represents endorse entity used for storage and retrieval.</param>
        /// <returns><see cref="Task"/> that represents endorse entity is saved or updated.</returns>
        Task<bool> StoreOrUpdateEndorsementDetailAsync(EndorsementEntity endorseEntity);

        /// <summary>
        /// Get award nomination endorsements from azure table storage.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <param name="awardCycleId">Active award cycle.</param>
        /// <param name="endorseeObjectId">Endorsee Azure Active Directory object id</param>
        /// <returns><see cref="Task"/>Returns endorse entity which is already saved.</returns>
        Task<IEnumerable<EndorsementEntity>> GetEndorsementsAsync(string teamId, string awardCycleId, string endorseeObjectId);
    }
}
