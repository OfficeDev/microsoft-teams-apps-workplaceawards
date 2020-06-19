// <copyright file="ConfigureAdminController.cs" company="Microsoft">
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
    /// This ASP controller is created to handle team captain requests and leverages TeamMemberUserPolicy for authorization.
    /// Dependency injection will provide the storage implementation <see cref="IConfigureAdminStorageProvider"/>
    /// <see cref="ITeamsInfoHelper"/> and logger implementation.
    /// Inherits <see cref="BaseRewardAndRecognitionController"/> to gather user claims for all incoming requests.
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    [Authorize(PolicyNames.MustBeTeamMemberUserPolicy)]
    public class ConfigureAdminController : BaseRewardAndRecognitionController
    {
        /// <summary>
        /// Provider to store logs in Azure Application Insights.
        /// </summary>
        private readonly ILogger<ConfigureAdminController> logger;

        /// <summary>
        /// Provider to fetch admin details from Azure Table Storage.
        /// </summary>
        private readonly IConfigureAdminStorageProvider storageProvider;

        /// <summary>
        /// Provider to fetch team details from bot adapter.
        /// </summary>
        private readonly ITeamsInfoHelper teamsInfoHelper;

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigureAdminController"/> class.
        /// </summary>
        /// <param name="logger">Instance to send logs to the application insights service.</param>
        /// <param name="storageProvider">Provider to store admin details in Azure Table Storage.</param>
        /// <param name="teamsInfoHelper">Provider to fetch team details from bot adapter.</param>
        public ConfigureAdminController(
            ILogger<ConfigureAdminController> logger,
            IConfigureAdminStorageProvider storageProvider,
            ITeamsInfoHelper teamsInfoHelper)
            : base()
        {
            this.logger = logger;
            this.storageProvider = storageProvider;
            this.teamsInfoHelper = teamsInfoHelper;
        }

        /// <summary>
        /// The Get API endpoint provides the list of team members based on specified team id.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <returns>List of members in team.</returns>
        [HttpGet("teammembers")]
        public async Task<IActionResult> GetTeamMembersAsync(string teamId)
        {
            try
            {
                if (teamId == null)
                {
                    return this.BadRequest(new { message = "Team ID cannot be empty." });
                }

                IEnumerable<TeamsChannelAccount> teamsChannelAccounts = new List<TeamsChannelAccount>();

                teamsChannelAccounts = await this.teamsInfoHelper.GetTeamMembersAsync(teamId);

                this.logger.LogInformation("GET call for fetching team members from team roster is successful.");
                return this.Ok(teamsChannelAccounts.Select(member => new { content = member.Email, header = member.Name, aadobjectid = member.AadObjectId }));
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error occurred while getting team member list.");
                throw;
            }
        }

        /// <summary>
        /// Post call to save team captain details in storage provider.
        /// </summary>
        /// <param name="adminDetails">Class contains details of team captain.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpPost("admindetail")]
        public async Task<IActionResult> SaveAdminDetailsAsync([FromBody]AdminEntity adminDetails)
        {
            try
            {
                if (adminDetails == null)
                {
                    return this.BadRequest(new { message = "Team captain detail entity cannot be null." });
                }

                // Validate admin must be a team member, and team member can only configure new admin.
                IEnumerable<TeamsChannelAccount> teamsChannelAccounts = new List<TeamsChannelAccount>();
                teamsChannelAccounts = await this.teamsInfoHelper.GetTeamMembersAsync(adminDetails.TeamId);
                if (!teamsChannelAccounts.Select(row => row.AadObjectId).Contains(adminDetails.CreatedByObjectId)
                    && teamsChannelAccounts.Select(row => row.AadObjectId).Contains(adminDetails.AdminObjectId))
                {
                    return this.BadRequest(new { message = "Invalid captain details, captain must be a team member." });
                }

                this.logger.LogInformation("Initiated call to on storage provider service.");
                var result = await this.storageProvider.UpsertAdminDetailAsync(adminDetails);
                this.logger.LogInformation("POST call for saving admin details in storage is successful.");
                return this.Ok(result);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while saving on admin details.");
                throw;
            }
        }

        /// <summary>
        /// This method returns team captain details for a given team id.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <returns>Admin entity</returns>
        [HttpGet("admindetail")]
        public async Task<IActionResult> GetAdminDetailsAsync(string teamId)
        {
            try
            {
                var adminDetails = await this.storageProvider.GetAdminDetailAsync(teamId);
                return this.Ok(adminDetails);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"failed to get admin details for team: {teamId}");
                throw;
            }
        }
    }
}
