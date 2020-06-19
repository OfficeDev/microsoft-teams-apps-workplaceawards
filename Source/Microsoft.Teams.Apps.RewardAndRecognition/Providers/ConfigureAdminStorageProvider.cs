// <copyright file="ConfigureAdminStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Providers
{
    using System;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Configure admin storage provider.
    /// </summary>
    public class ConfigureAdminStorageProvider : StorageBaseProvider, IConfigureAdminStorageProvider
    {
        /// <summary>
        /// Admin detail table.
        /// </summary>
        public const string AdminTable = "AdminDetail";

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigureAdminStorageProvider"/> class.
        /// </summary>
        /// <param name="storageOptions">A set of key/value application storage configuration properties.</param>
        public ConfigureAdminStorageProvider(IOptions<StorageOptions> storageOptions)
        : base(storageOptions?.Value, AdminTable)
        {
            if (storageOptions == null)
            {
                throw new ArgumentNullException(nameof(storageOptions));
            }
        }

        /// <summary>
        /// Save admin details in Azure table storage.
        /// </summary>
        /// <param name="adminDetails">Admin details to be stored in Azure table storage.</param>
        /// <returns><see cref="Task"/> Returns admin entity after save.</returns>
        public async Task<AdminEntity> UpsertAdminDetailAsync(AdminEntity adminDetails)
        {
            await this.EnsureInitializedAsync();
            adminDetails = adminDetails ?? throw new ArgumentNullException(nameof(adminDetails));
            adminDetails.RowUniqueId = Guid.NewGuid().ToString();
            TableOperation addOperation = TableOperation.Insert(adminDetails);
            var result = await this.CloudTable.ExecuteAsync(addOperation);
            return result.Result as AdminEntity;
        }

        /// <summary>
        /// Get already saved admin detail from storage table for given team id.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <returns><see cref="Task"/>Returns admin entity.</returns>
        public async Task<AdminEntity> GetAdminDetailAsync(string teamId)
        {
            await this.EnsureInitializedAsync();
            AdminEntity adminDetails;
            var query = new TableQuery<AdminEntity>().Where(TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, teamId));
            TableContinuationToken tableContinuationToken = null;

            do
            {
                var queryResponse = await this.CloudTable.ExecuteQuerySegmentedAsync(query, tableContinuationToken);
                tableContinuationToken = queryResponse.ContinuationToken;
                adminDetails = queryResponse.Results.OrderByDescending(rows => rows.CreatedOn).FirstOrDefault();
            }
            while (tableContinuationToken != null);

            return adminDetails;
        }
    }
}