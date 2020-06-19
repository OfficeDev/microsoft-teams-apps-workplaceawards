// <copyright file="EndorsementsStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Providers
{
    using System;
    using System.Collections.Generic;
    using System.Net;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Endorse storage provider.
    /// </summary>
    public class EndorsementsStorageProvider : StorageBaseProvider, IEndorsementsStorageProvider
    {
        /// <summary>
        /// endorsement table.
        /// </summary>
        private const string EndorseTable = "EndorseDetail";

        /// <summary>
        /// Initializes a new instance of the <see cref="EndorsementsStorageProvider"/> class.
        /// </summary>
        /// <param name="storageOptions">A set of key/value application storage configuration properties.</param>
        public EndorsementsStorageProvider(IOptions<StorageOptions> storageOptions)
            : base(storageOptions?.Value, EndorseTable)
        {
            if (storageOptions == null)
            {
                throw new ArgumentNullException(nameof(storageOptions));
            }
        }

        /// <summary>
        /// Store or update endorsement in Azure table storage.
        /// </summary>
        /// <param name="endorseEntity">Represents endorse entity used for storage and retrieval.</param>
        /// <returns><see cref="Task"/> that represents endorse entity is saved or updated.</returns>
        public async Task<bool> StoreOrUpdateEndorsementDetailAsync(EndorsementEntity endorseEntity)
        {
            await this.EnsureInitializedAsync();
            endorseEntity = endorseEntity ?? throw new ArgumentNullException(nameof(endorseEntity));
            endorseEntity.RowUniqueId = Guid.NewGuid().ToString();
            TableOperation addOrUpdateOperation = TableOperation.InsertOrReplace(endorseEntity);
            var result = await this.CloudTable.ExecuteAsync(addOrUpdateOperation);
            return result.HttpStatusCode == (int)HttpStatusCode.NoContent;
        }

        /// <summary>
        /// Get award nomination endorsements from azure table storage.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <param name="awardCycleId">Active award cycle.</param>
        /// <param name="endorseeObjectId">Endorsee Azure Active Directory object id.</param>
        /// <returns><see cref="Task"/> Already endorsement.</returns>
        public async Task<IEnumerable<EndorsementEntity>> GetEndorsementsAsync(string teamId, string awardCycleId, string endorseeObjectId)
        {
            await this.EnsureInitializedAsync();

            var endorseEntity = new List<EndorsementEntity>();
            string partitionKeyCondition = TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, teamId);
            string awardCycleIdCondition = TableQuery.GenerateFilterCondition("AwardCycle", QueryComparisons.Equal, awardCycleId);
            string endorseeObjectIdCondition = TableQuery.GenerateFilterCondition("EndorseeObjectId", QueryComparisons.Equal, endorseeObjectId);
            string condition = TableQuery.CombineFilters(partitionKeyCondition, TableOperators.And, awardCycleIdCondition);

            // endorseeObjectId: is optional parameter to get individual endorsement details.
            if (!string.IsNullOrWhiteSpace(endorseeObjectId))
            {
                condition = TableQuery.CombineFilters(condition, TableOperators.And, endorseeObjectIdCondition);
            }

            TableQuery<EndorsementEntity> query = new TableQuery<EndorsementEntity>().Where(condition);
            TableContinuationToken tableContinuationToken = null;

            do
            {
                var queryResponse = await this.CloudTable.ExecuteQuerySegmentedAsync(query, tableContinuationToken);
                tableContinuationToken = queryResponse.ContinuationToken;
                endorseEntity.AddRange(queryResponse?.Results);
            }
            while (tableContinuationToken != null);

            return endorseEntity;
        }
    }
}