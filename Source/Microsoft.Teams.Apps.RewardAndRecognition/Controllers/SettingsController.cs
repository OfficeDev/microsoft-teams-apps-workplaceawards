// <copyright file="SettingsController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Controllers
{
    using System;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.RewardAndRecognition.Authentication.AuthenticationPolicy;

    /// <summary>
    /// This ASP controller is created to handle award requests and leverages TeamMemberUserPolicy for authorization.
    /// Dependency injection will provide the storage implementation and logger.
    /// Inherits <see cref="BaseRewardAndRecognitionController"/> to gather user claims for all incoming requests.
    /// The class provides endpoint to share required application settings.
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    [Authorize(PolicyNames.MustBeTeamMemberUserPolicy)]
    public class SettingsController : BaseRewardAndRecognitionController
    {
        /// <summary>
        /// Provider to store logs in Azure Application Insights.
        /// </summary>
        private readonly ILogger<SettingsController> logger;

        /// <summary>
        /// Represents a set of key/value application configuration properties.
        /// </summary>
        private readonly IConfiguration configuration;

        /// <summary>
        /// Initializes a new instance of the <see cref="SettingsController"/> class.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="configuration">configuration.</param>
        public SettingsController(ILogger<SettingsController> logger, IConfiguration configuration)
        {
            this.logger = logger;
            this.configuration = configuration;
        }

        /// <summary>
        /// Get bot setting to client application.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <returns>Bot id.</returns>
        [HttpGet("botsettings")]
        public IActionResult GetBotSettings(string teamId)
        {
            try
            {
                return this.Ok(new
                {
                    botId = this.configuration["MicrosoftAppId"],
                    instrumentationKey = this.configuration["ApplicationInsights:InstrumentationKey"],
                });
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error while fetching bot setting for team: {teamId}");
                throw;
            }
        }
    }
}