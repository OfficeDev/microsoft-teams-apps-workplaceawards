// <copyright file="NominationsController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.RewardAndRecognition.Authentication.AuthenticationPolicy;
    using Microsoft.Teams.Apps.RewardAndRecognition.Helpers;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;
    using Microsoft.Teams.Apps.RewardAndRecognition.Providers;

    /// <summary>
    /// This ASP controller is created to handle award requests and leverages TeamMemberUserPolicy for authorization.
    /// Dependency injection will provide the storage implementation and logger.
    /// Inherits <see cref="BaseRewardAndRecognitionController"/> to gather user claims for all incoming requests.
    /// The class provides endpoint to handle award nominations requests.
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    [Authorize(PolicyNames.MustBeTeamMemberUserPolicy)]
    public class NominationsController : BaseRewardAndRecognitionController
    {
        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<NominationsController> logger;

        /// <summary>
        /// Provider to search nomination details in Azure search service.
        /// </summary>
        private readonly IAwardNominationStorageProvider storageProvider;

        /// <summary>
        /// Provider for fetching information about endorsement details from storage table.
        /// </summary>
        private readonly IEndorsementsStorageProvider endorseStorageProvider;

        /// <summary>
        /// Provider to fetch team details from bot adapter.
        /// </summary>
        private readonly ITeamsInfoHelper teamsInfoHelper;

        /// <summary>
        /// Initializes a new instance of the <see cref="NominationsController"/> class.
        /// </summary>
        /// <param name="logger">Provider to store logs in Azure Application Insights.</param>
        /// <param name="storageProvider">Nominate award detail storage provider.</param>
        /// <param name="endorseStorageProvider">Endorsement storage provider.</param>
        /// <param name="teamsInfoHelper">Provider to fetch team details from bot adapter.</param>
        public NominationsController(
            ILogger<NominationsController> logger,
            IAwardNominationStorageProvider storageProvider,
            IEndorsementsStorageProvider endorseStorageProvider,
            ITeamsInfoHelper teamsInfoHelper)
        {
            this.logger = logger;
            this.storageProvider = storageProvider;
            this.endorseStorageProvider = endorseStorageProvider;
            this.teamsInfoHelper = teamsInfoHelper;
        }

        /// <summary>
        /// Post call to save nominated award details in Azure Table storage.
        /// </summary>
        /// <param name="nominateDetails">Class contains details of on award nomination.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpPost("nomination")]
        public async Task<IActionResult> SaveAwardNominationAsync([FromBody]NominationEntity nominateDetails)
        {
            try
            {
                if (nominateDetails == null)
                {
                    return this.BadRequest(new { message = "Nomination details can not be null." });
                }

                var userClaim = this.GetUserClaims();

                // Validate nominee and nominator belongs to same team.
                IEnumerable<TeamsChannelAccount> teamsChannelAccounts = new List<TeamsChannelAccount>();
                teamsChannelAccounts = await this.teamsInfoHelper.GetTeamMembersAsync(nominateDetails.TeamId);
                var nominees = nominateDetails.NomineeObjectIds.Split(",").ToList();
                if (!(nominees.TrueForAll(nomineeAadObjectId => teamsChannelAccounts.Select(row => row.AadObjectId).Contains(nomineeAadObjectId.Trim()))
                    && teamsChannelAccounts.Select(row => row.AadObjectId).Contains(nominateDetails.NominatedByObjectId)
                    && userClaim.FromId == nominateDetails.NominatedByObjectId))
                {
                    return this.BadRequest(new { message = "Invalid nomination details, nominee and nominator must be from a same team." });
                }

                // Check for duplicate award nomination.
                var isAlreadyNominated = await this.storageProvider.CheckDuplicateNominationAsync(nominateDetails.TeamId, nominateDetails.NomineeObjectIds, nominateDetails.RewardCycleId, nominateDetails.AwardId, nominateDetails.NominatedByObjectId);
                if (isAlreadyNominated)
                {
                    return this.BadRequest(new { message = "You have already nominated selected team member(s) for this award. Feel free to nominate others!" });
                }

                this.logger.LogInformation("Initiated call to on storage provider service.");
                var result = await this.storageProvider.StoreOrUpdateAwardNominationAsync(nominateDetails);
                this.logger.LogInformation("POST call for nominated award details in storage is successful.");
                return this.Ok(result);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while saving nominated award details.");
                throw;
            }
        }

        /// <summary>
        /// This method checks the duplication nomination requests for a given award.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <param name="aadObjectIds">Comma separated Azure active directory object Id of nominees.</param>
        /// <param name="cycleId">Active reward cycle id.</param>
        /// <param name="awardId">Award unique id.</param>
        /// <param name="nominatedByObjectId">Azure active directory object Id of nominator.</param>
        /// <returns>Returns true if same group of user already nominated, else return false.</returns>
        [HttpGet("checkduplicatenomination")]
        public async Task<IActionResult> CheckDuplicateNominationAsync(string teamId, string aadObjectIds, string cycleId, string awardId, string nominatedByObjectId)
        {
            try
            {
                var isAlreadyNominated = await this.storageProvider.CheckDuplicateNominationAsync(teamId, aadObjectIds, cycleId, awardId, nominatedByObjectId);
                return this.Ok(isAlreadyNominated);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"failed to get nomination details for team: {aadObjectIds}");
                throw;
            }
        }

        /// <summary>
        /// This method returns all nominations for a given team Id.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <param name="isAwardGranted">True for published awards, else false.</param>
        /// <param name="awardCycleId">Active award cycle.</param>
        /// <returns>Returns all nominations</returns>
        [HttpGet("allnominations")]
        public async Task<IActionResult> GetAwardNominationAsync(string teamId, bool isAwardGranted, string awardCycleId)
        {
            try
            {
                var publishAwardDetails = new List<PublishResult>();
                var nominations = await this.storageProvider.GetAwardNominationAsync(teamId, isAwardGranted, awardCycleId);
                var endorseDetails = await this.endorseStorageProvider.GetEndorsementsAsync(teamId, awardCycleId, endorseeObjectId: string.Empty);
                publishAwardDetails = nominations.Select(nomination => new PublishResult()
                {
                    AwardCycle = string.Empty,
                    AwardName = nomination.AwardName,
                    NominatedByName = nomination.NominatedByName,
                    AwardId = nomination.AwardId,
                    NominatedByObjectId = nomination.NominatedByObjectId,
                    NominationId = nomination.NominationId,
                    NominatedByUserPrincipalName = nomination.NominatedByUserPrincipalName,
                    NomineeNames = nomination.NomineeNames,
                    GroupName = nomination.GroupName,
                    NomineeObjectIds = nomination.NomineeObjectIds,
                    NomineeUserPrincipalNames = nomination.NomineeUserPrincipalNames,
                    RewardCycleId = nomination.RewardCycleId,
                    NominatedOn = nomination.NominatedOn,
                    ReasonForNomination = nomination.ReasonForNomination,
                    EndorsementCount = endorseDetails.Where(agg => agg.EndorseeObjectId == nomination.NomineeObjectIds && nomination.AwardId == agg.EndorsedForAwardId).Count(),
                }).OrderByDescending(row => row.NominatedOn).ToList();

                return this.Ok(publishAwardDetails);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"failed to get nominations for team: {teamId}");
                throw;
            }
        }

        /// <summary>
        /// This method publish award nominations.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <param name="nominationIds">Comma seperated nomination ids.</param>
        /// <returns>Returns true if publish succeeded</returns>
        [HttpGet("publishnominations")]
        [Authorize(PolicyNames.MustBeTeamCaptainUserPolicy)]
        public async Task<IActionResult> UpdateAwardNominationAsync(string teamId, string nominationIds)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(teamId))
                {
                    return this.BadRequest(new { message = "Team id can not be null or empty." });
                }

                if (string.IsNullOrWhiteSpace(nominationIds))
                {
                    return this.BadRequest(new { message = "Nomination ids can not be null or empty." });
                }

                var result = await this.storageProvider.PublishAwardNominationAsync(teamId, nominationIds.Split(','));
                return this.Ok(result);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"failed to publish nominations for team: {teamId}");
                throw;
            }
        }
    }
}