// <copyright file="RewardCycleController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Controllers
{
    using System;
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
    /// The class provides endpoint to manage reward cycles.
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    [Authorize(PolicyNames.MustBeTeamMemberUserPolicy)]
    public class RewardCycleController : BaseRewardAndRecognitionController
    {
        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<RewardCycleController> logger;

        /// <summary>
        /// Provider for fetching information about active award cycle details from storage table.
        /// </summary>
        private readonly IRewardCycleStorageProvider storageProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="RewardCycleController"/> class.
        /// </summary>
        /// <param name="logger">Provider to store logs in Azure Application Insights.</param>
        /// <param name="storageProvider">Reward cycle storage provider.</param>
        public RewardCycleController(ILogger<RewardCycleController> logger, IRewardCycleStorageProvider storageProvider)
        {
            this.logger = logger;
            this.storageProvider = storageProvider;
        }

        /// <summary>
        /// This method returns reward cycle information for a given team.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <param name="isActiveCycle">Reward cycle state.</param>
        /// <returns>Current reward cycle if active else returns last published reward cycle details.</returns>
        [HttpGet("rewardcycledetails")]
        public async Task<IActionResult> GetRewardCycleAsync(string teamId, bool isActiveCycle = true)
        {
            RewardCycleEntity rewardCycle;
            try
            {
                if (isActiveCycle)
                {
                    rewardCycle = await this.storageProvider.GetCurrentRewardCycleAsync(teamId);
                }
                else
                {
                    rewardCycle = await this.storageProvider.GetPublishedRewardCycleAsync(teamId);
                }

                return this.Ok(rewardCycle);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"failed to get awards for team: {teamId}");
                throw;
            }
        }

        /// <summary>
        /// Post call to store reward cycle details by Team captain authorized user.
        /// </summary>
        /// <param name="rewardCycleEntity">Holds reward cycle detail entity data.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpPost("rewardcycle")]
        [Authorize(PolicyNames.MustBeTeamCaptainUserPolicy)]
        public async Task<IActionResult> PostAsync([FromBody] RewardCycleEntity rewardCycleEntity)
        {
            try
            {
                if (rewardCycleEntity == null)
                {
                    this.logger.LogInformation("Set reward cycle entity is null.");
                    return this.BadRequest(new { message = "Award cycle entity can not be null." });
                }

                if (rewardCycleEntity.RewardCycleStartDate == null)
                {
                    this.logger.LogInformation("set reward cycle start date is null.");
                    return this.BadRequest(new { message = "Award cycle start date can not be null." });
                }

                if (rewardCycleEntity.RewardCycleEndDate == null)
                {
                    this.logger.LogInformation("Award cycle end date is null.");
                    return this.BadRequest(new { message = "Award cycle end date can not be null." });
                }

                if (rewardCycleEntity.CycleId == null)
                {
                    rewardCycleEntity.CycleId = Guid.NewGuid().ToString();
                    rewardCycleEntity.CreatedOn = DateTime.UtcNow;
                }

                return this.Ok(await this.storageProvider.StoreOrUpdateRewardCycleAsync(rewardCycleEntity));
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to award service.");
                throw;
            }
        }
    }
}