// <copyright file="IConfigureAdminStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Providers
{
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;

    /// <summary>
    /// Interface for configure admin storage provider.
    /// </summary>
    public interface IConfigureAdminStorageProvider
    {
        /// <summary>
        /// Save admin details in Azure table storage.
        /// </summary>
        /// <param name="adminDetails">Admin details to be stored in table storage.</param>
        /// <returns><see cref="Task"/>Returns admin entity when saved successfully.</returns>
        Task<AdminEntity> UpsertAdminDetailAsync(AdminEntity adminDetails);

        /// <summary>
        /// Get already saved admin entity detail from storage table.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <returns><see cref="Task"/>Returns already saved admin entity detail.</returns>
        Task<AdminEntity> GetAdminDetailAsync(string teamId);
    }
}