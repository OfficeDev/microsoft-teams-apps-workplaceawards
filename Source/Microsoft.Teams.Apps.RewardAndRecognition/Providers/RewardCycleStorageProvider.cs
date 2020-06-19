// <copyright file="RewardCycleStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Providers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Reward cycle storage provider class.
    /// </summary>
    public class RewardCycleStorageProvider : StorageBaseProvider, IRewardCycleStorageProvider
    {
        private const string RewardCycleTable = "RewardCycleDetail";

        /// <summary>
        /// Provider to store logs in Azure Application Insights.
        /// </summary>
        private readonly ILogger<RewardCycleStorageProvider> logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="RewardCycleStorageProvider"/> class.
        /// </summary>
        /// <param name="storageOptions">A set of key/value application storage configuration properties.</param>
        /// <param name="logger">Instance to send logs to the application insights service.</param>
        public RewardCycleStorageProvider(IOptions<StorageOptions> storageOptions, ILogger<RewardCycleStorageProvider> logger)
            : base(storageOptions?.Value, RewardCycleTable)
        {
            if (storageOptions == null)
            {
                throw new ArgumentNullException(nameof(storageOptions));
            }

            this.logger = logger;
        }

        /// <summary>
        /// This method is used to fetch active/unpublished reward cycle details for a given team Id.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <returns>Reward cycle for a given team Id.</returns>
        public async Task<RewardCycleEntity> GetCurrentRewardCycleAsync(string teamId)
        {
            await this.EnsureInitializedAsync();
            string filterTeamId = TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, teamId);
            string filterActiveCycle = TableQuery.GenerateFilterConditionForInt("RewardCycleState", QueryComparisons.Equal, (int)RewardCycleState.Active);
            string filterInActiveCycle = TableQuery.GenerateFilterConditionForInt("RewardCycleState", QueryComparisons.Equal, (int)RewardCycleState.Inactive);
            string filterPublish = TableQuery.GenerateFilterConditionForInt("ResultPublished", QueryComparisons.Equal, (int)ResultPublishState.Unpublished);
            string combineFilter = TableQuery.CombineFilters(filterInActiveCycle, TableOperators.And, filterPublish);
            string filter = TableQuery.CombineFilters(filterTeamId, TableOperators.And, TableQuery.CombineFilters(filterActiveCycle, TableOperators.Or, combineFilter));
            var query = new TableQuery<RewardCycleEntity>().Where(filter);
            TableContinuationToken continuationToken = null;
            var cycles = new List<RewardCycleEntity>();

            do
            {
                var queryResult = await this.CloudTable.ExecuteQuerySegmentedAsync(query, continuationToken);
                cycles.AddRange(queryResult?.Results);
                continuationToken = queryResult?.ContinuationToken;
            }
            while (continuationToken != null);

            return cycles.FirstOrDefault();
        }

        /// <summary>
        /// This method is used to fetch published reward cycle details for a given team Id.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <returns>Reward cycle for a given team Id.</returns>
        public async Task<RewardCycleEntity> GetPublishedRewardCycleAsync(string teamId)
        {
            await this.EnsureInitializedAsync();
            string filterTeamId = TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, teamId);
            string filterPublish = TableQuery.GenerateFilterConditionForInt("ResultPublished", QueryComparisons.Equal, (int)ResultPublishState.Published);
            string filter = TableQuery.CombineFilters(filterTeamId, TableOperators.And, filterPublish);
            var query = new TableQuery<RewardCycleEntity>().Where(filter);
            TableContinuationToken continuationToken = null;
            var cycles = new List<RewardCycleEntity>();

            do
            {
                var queryResult = await this.CloudTable.ExecuteQuerySegmentedAsync(query, continuationToken);
                cycles.AddRange(queryResult?.Results);
                continuationToken = queryResult?.ContinuationToken;
            }
            while (continuationToken != null);

            return cycles.OrderByDescending(row => row.ResultPublishedOn).FirstOrDefault();
        }

        /// <summary>
        /// Store or update reward cycle in table storage.
        /// </summary>
        /// <param name="rewardCycleEntity">Represents reward cycle entity used for storage and retrieval.</param>
        /// <returns><see cref="Task"/> that represents reward cycle entity is saved or updated.</returns>
        public async Task<RewardCycleEntity> StoreOrUpdateRewardCycleAsync(RewardCycleEntity rewardCycleEntity)
        {
            await this.EnsureInitializedAsync();
            TableOperation addOrUpdateOperation = TableOperation.InsertOrReplace(rewardCycleEntity);
            var result = await this.CloudTable.ExecuteAsync(addOrUpdateOperation);
            return result.Result as RewardCycleEntity;
        }

        /// <summary>
        /// This method is used get active reward cycle details for all teams.
        /// </summary>
        /// <returns><see cref="Task"/> that represents reward cycle entity is saved or updated.</returns>
        public async Task<List<RewardCycleEntity>> GetActiveRewardCycleForAllTeamsAsync()
        {
            await this.EnsureInitializedAsync();

            // Get all active reward cycle
            string filterActiveCycle = TableQuery.GenerateFilterConditionForInt("RewardCycleState", QueryComparisons.Equal, (int)RewardCycleState.Active);
            var query = new TableQuery<RewardCycleEntity>().Where(filterActiveCycle);
            TableContinuationToken continuationToken = null;
            var activeCycles = new List<RewardCycleEntity>();

            do
            {
                var queryResult = await this.CloudTable.ExecuteQuerySegmentedAsync(query, continuationToken);
                activeCycles.AddRange(queryResult?.Results);
                continuationToken = queryResult?.ContinuationToken;
            }
            while (continuationToken != null);
            return activeCycles as List<RewardCycleEntity>;
        }

        /// <summary>
        /// This method is used to fetch current reward cycle for all teams.
        /// </summary>
        /// <returns><see cref="Task"/> that represents reward cycle entity is saved or updated.</returns>
        public async Task<List<RewardCycleEntity>> GetCurrentRewardCycleForAllTeamsAsync()
        {
            await this.EnsureInitializedAsync();

            var query = new TableQuery<RewardCycleEntity>();
            TableContinuationToken continuationToken = null;
            var currentRewardCycles = new List<RewardCycleEntity>();

            do
            {
                var queryResult = await this.CloudTable.ExecuteQuerySegmentedAsync(query, continuationToken);
                currentRewardCycles.AddRange(queryResult?.Results);
                continuationToken = queryResult?.ContinuationToken;
            }
            while (continuationToken != null);

            currentRewardCycles = currentRewardCycles.GroupBy(row => row.TeamId, (key, group) => group.OrderByDescending(rewardCycle => rewardCycle.Timestamp).FirstOrDefault()).ToList();

            return currentRewardCycles;
        }
    }
}