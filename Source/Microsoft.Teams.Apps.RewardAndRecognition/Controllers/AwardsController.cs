// <copyright file="AwardsController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.RewardAndRecognition.Authentication.AuthenticationPolicy;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;
    using Microsoft.Teams.Apps.RewardAndRecognition.Providers;

    /// <summary>
    /// This ASP controller is created to handle award requests and leverages TeamMemberUserPolicy for authorization.
    /// Dependency injection will provide the storage implementation and logger.
    /// Inherits <see cref="BaseRewardAndRecognitionController"/> to gather user claims for all incoming requests.
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    [Authorize(PolicyNames.MustBeTeamMemberUserPolicy)]
    public class AwardsController : BaseRewardAndRecognitionController
    {
        /// <summary>
        /// Instance to send logs to the application insights service.
        /// </summary>
        private readonly ILogger<AwardsController> logger;

        /// <summary>
        /// Helper for fetching award details from azure table storage.
        /// </summary>
        private readonly IAwardsStorageProvider storageProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="AwardsController"/> class.
        /// </summary>
        /// <param name="logger">Provider to store logs in Azure Application Insights.</param>
        /// <param name="storageProvider">Awards storage provider.</param>
        public AwardsController(ILogger<AwardsController> logger, IAwardsStorageProvider storageProvider)
        {
            this.logger = logger;
            this.storageProvider = storageProvider;
        }

        /// <summary>
        /// This method returns all awards for a given team Id.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <returns>Awards</returns>
        [HttpGet("allawards")]
        public async Task<IActionResult> GetAwardsAsync(string teamId)
        {
            try
            {
                var awards = await this.storageProvider.GetAwardsAsync(teamId);
                return this.Ok(awards);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"failed to get awards for team: {teamId}");
                throw;
            }
        }

        /// <summary>
        /// This method returns award details for a given team Id and awardId.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <param name="awardId">Award Id.</param>
        /// <returns>Award Details</returns>
        [HttpGet("awarddetails")]
        public async Task<IActionResult> GetAwardDetailsAsync(string teamId, string awardId)
        {
            try
            {
                var award = await this.storageProvider.GetAwardDetailsAsync(teamId, awardId);
                return this.Ok(award);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"failed to get award details for team: {teamId}");
                throw;
            }
        }

        /// <summary>
        /// Post call to store award details data in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="awardEntity">Holds award detail entity data.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpPost("award")]
        [Authorize(PolicyNames.MustBeTeamCaptainUserPolicy)]
        public async Task<IActionResult> PostAsync([FromBody] AwardEntity awardEntity)
        {
            try
            {
                if (awardEntity == null)
                {
                    this.logger.LogError("Award entity is null.");
                    return this.BadRequest(new { message = "Award entity can not be null." });
                }

                if (string.IsNullOrEmpty(awardEntity.AwardName))
                {
                    this.logger.LogError("Award name is empty.");
                    return this.BadRequest(new { message = "Award name can not be empty." });
                }

                if (string.IsNullOrEmpty(awardEntity.AwardDescription))
                {
                    this.logger.LogError("Award description is empty.");
                    return this.BadRequest(new { message = "Award description can not be empty." });
                }

                var claims = this.GetUserClaims();
                this.logger.LogInformation("Adding award");
                if (awardEntity.AwardId == null)
                {
                    awardEntity.AwardId = Guid.NewGuid().ToString();
                    awardEntity.CreatedOn = DateTime.UtcNow;
                    awardEntity.CreatedBy = claims.FromId;
                }
                else
                {
                    awardEntity.ModifiedBy = claims.FromId;
                }

                return this.Ok(await this.storageProvider.StoreOrUpdateAwardAsync(awardEntity));
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to award service.");
                throw;
            }
        }

        /// <summary>
        /// Delete call to delete award details data in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="teamId">Holds team Id.</param>
        /// <param name="awardIds">User selected response Ids.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpDelete("awards")]
        [Authorize(PolicyNames.MustBeTeamCaptainUserPolicy)]
        public async Task<IActionResult> DeleteAsync(string teamId, string awardIds)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(teamId))
                {
                    this.logger.LogError("Error while deleting award details data in Microsoft Azure Table storage. Team id is null");
                    return this.BadRequest(new { message = "Team Id can not be null or empty." });
                }

                if (string.IsNullOrWhiteSpace(awardIds))
                {
                    this.logger.LogError("Error while deleting award details data in Microsoft Azure Table storage. Award id is null");
                    return this.BadRequest(new { message = "Award Ids can not be null or empty." });
                }

                IList<string> awards = awardIds.Split(",");
                return this.Ok(await this.storageProvider.DeleteAwardsAsync(teamId, awards));
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while deleting awards");
                throw;
            }
        }
    }
}