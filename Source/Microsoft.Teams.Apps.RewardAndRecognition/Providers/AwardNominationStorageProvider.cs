// <copyright file="AwardNominationStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Providers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Nominate award storage provider.
    /// </summary>
    public class AwardNominationStorageProvider : StorageBaseProvider, IAwardNominationStorageProvider
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="AwardNominationStorageProvider"/> class.
        /// </summary>
        /// <param name="storageOptions">A set of key/value application storage configuration properties.</param>
        public AwardNominationStorageProvider(IOptions<StorageOptions> storageOptions)
            : base(storageOptions?.Value, Constants.NominateAwardTable)
        {
            if (storageOptions == null)
            {
                throw new ArgumentNullException(nameof(storageOptions));
            }
        }

        /// <summary>
        /// Get already nominate entity detail from storage table.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <returns><see cref="Task"/> Already saved nominate entity detail.</returns>
        public async Task<NominationEntity> GetAwardNominationAsync(string teamId)
        {
            await this.EnsureInitializedAsync();
            if (string.IsNullOrEmpty(teamId))
            {
                return null;
            }

            var searchOperation = TableOperation.Retrieve<NominationEntity>("PartitionKey", teamId);
            var searchResult = await this.CloudTable.ExecuteAsync(searchOperation);

            return (NominationEntity)searchResult.Result;
        }

        /// <summary>
        /// This method is used to check duplicate nominations.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <param name="nomineeObjectIds">Comma separated Azure active directory object Id of nominees.</param>
        /// <param name="awardCycleId">Active award cycle.</param>
        /// <param name="awardId">Award unique id.</param>
        /// <param name="nominatedByObjectId">Azure active directory object Id of nominator.</param>
        /// <returns>Returns true if same group of user already nominated, else return false.</returns>
        public async Task<bool> CheckDuplicateNominationAsync(string teamId, string nomineeObjectIds, string awardCycleId, string awardId, string nominatedByObjectId)
        {
            await this.EnsureInitializedAsync();
            var nominateEntity = new List<NominationEntity>();
            string partitionKeyCondition = TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, teamId);
            string activeCycleCondition = TableQuery.GenerateFilterCondition("RewardCycleId", QueryComparisons.Equal, awardCycleId);
            string awardCondition = TableQuery.GenerateFilterCondition("AwardId", QueryComparisons.Equal, awardId);
            string nominatedByCondition = TableQuery.GenerateFilterCondition("NominatedByObjectId", QueryComparisons.Equal, nominatedByObjectId);
            string condition = TableQuery.CombineFilters(partitionKeyCondition, TableOperators.And, awardCondition);
            condition = TableQuery.CombineFilters(condition, TableOperators.And, activeCycleCondition);
            condition = TableQuery.CombineFilters(condition, TableOperators.And, nominatedByCondition);
            TableQuery<NominationEntity> query = new TableQuery<NominationEntity>().Where(condition);
            TableContinuationToken tableContinuationToken = null;

            do
            {
                var queryResponse = await this.CloudTable.ExecuteQuerySegmentedAsync(query, tableContinuationToken);
                tableContinuationToken = queryResponse.ContinuationToken;
                nominateEntity.AddRange(queryResponse?.Results);
            }
            while (tableContinuationToken != null);

            return this.CheckDuplicateNomination(nominateEntity, nomineeObjectIds);
        }

        /// <summary>
        /// Store or update nominated details in Azure table storage.
        /// </summary>
        /// <param name="nominateEntity">Represents nominate entity used for storage and retrieval.</param>
        /// <returns><see cref="Task"/>Returns nominate entity.</returns>
        public async Task<NominationEntity> StoreOrUpdateAwardNominationAsync(NominationEntity nominateEntity)
        {
            await this.EnsureInitializedAsync();
            nominateEntity = nominateEntity ?? throw new ArgumentNullException(nameof(nominateEntity));
            nominateEntity.NominationId = Guid.NewGuid().ToString();
            TableOperation addOrUpdateOperation = TableOperation.InsertOrReplace(nominateEntity);
            var result = await this.CloudTable.ExecuteAsync(addOrUpdateOperation);
            return result.Result as NominationEntity;
        }

        /// <summary>
        /// This method is used to fetch award details for a given team Id and awardId.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <param name="isAwardGranted">True if award granted, else false.</param>
        /// <param name="awardCycleId">Active award cycle.</param>
        /// <returns>Nomination details.</returns>
        public async Task<IEnumerable<NominationEntity>> GetAwardNominationAsync(string teamId, bool isAwardGranted, string awardCycleId)
        {
            await this.EnsureInitializedAsync();

            var nominateEntity = new List<NominationEntity>();
            string partitionKeyCondition = TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, teamId);
            string awardGrantedCondition = TableQuery.GenerateFilterConditionForBool("AwardGranted", QueryComparisons.Equal, isAwardGranted);
            string activeCycleCondition = TableQuery.GenerateFilterCondition("RewardCycleId", QueryComparisons.Equal, awardCycleId);
            string condition = TableQuery.CombineFilters(partitionKeyCondition, TableOperators.And, awardGrantedCondition);
            condition = TableQuery.CombineFilters(condition, TableOperators.And, activeCycleCondition);
            TableQuery<NominationEntity> query = new TableQuery<NominationEntity>().Where(condition);
            TableContinuationToken tableContinuationToken = null;

            do
            {
                var queryResponse = await this.CloudTable.ExecuteQuerySegmentedAsync(query, tableContinuationToken);
                tableContinuationToken = queryResponse.ContinuationToken;
                nominateEntity.AddRange(queryResponse?.Results);
            }
            while (tableContinuationToken != null);
            return nominateEntity as List<NominationEntity>;
        }

        /// <summary>
        /// This method is used to publish nomination details for a given team Id.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <param name="nominationIds">Published nomination ids.</param>
        /// <returns>Nomination details.</returns>
        public async Task<bool> PublishAwardNominationAsync(string teamId, IEnumerable<string> nominationIds)
        {
            await this.EnsureInitializedAsync();
            if (nominationIds == null)
            {
                throw new ArgumentNullException(nameof(nominationIds));
            }

            foreach (var nominationId in nominationIds)
            {
                var operation = TableOperation.Retrieve<NominationEntity>(teamId, nominationId);
                var data = await this.CloudTable.ExecuteAsync(operation);
                var entity = data.Result as NominationEntity;

                entity.AwardGranted = true;
                entity.AwardPublishedOn = DateTime.UtcNow;
                TableOperation updateOperation = TableOperation.InsertOrReplace(entity);
                var result = await this.CloudTable.ExecuteAsync(updateOperation);
            }

            return true;
        }

        /// <summary>
        /// Check for duplicate award nomination.
        /// </summary>
        /// <param name="existingNominations">Existing nominations.</param>
        /// <param name="newNomination">New nominations.</param>
        /// <returns>Returns true if same group of user already nominated, else return false.</returns>
        private bool CheckDuplicateNomination(List<NominationEntity> existingNominations, string newNomination)
        {
            bool isAlreadyNominated = false;
            foreach (var nominees in existingNominations.Select(row => row.NomineeObjectIds))
            {
                if (nominees == newNomination)
                {
                    return true;
                }
                else
                {
                    IEnumerable<string> existingNominees = nominees.Split(',').Select(a => a.Trim());
                    IEnumerable<string> newNominees = newNomination.Split(',').Select(a => a.Trim());
                    if (!existingNominees.Except(newNominees).Any() && !newNominees.Except(existingNominees).Any())
                    {
                        isAlreadyNominated = true;
                        break;
                    }
                }
            }

            return isAlreadyNominated;
        }
    }
}